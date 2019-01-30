VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{DBF30C82-CAF3-11D5-84FF-0050BA3D926D}#8.5#0"; "VLMnuPlus.ocx"
Begin VB.Form frmMain 
   Caption         =   "DHTML Menu Builder - [untitled]"
   ClientHeight    =   10155
   ClientLeft      =   3210
   ClientTop       =   3345
   ClientWidth     =   14160
   FillColor       =   &H00DEE3E7&
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   10155
   ScaleWidth      =   14160
   Begin VB.Timer tmrResetLivePreviewBusyState 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5730
      Top             =   9345
   End
   Begin SmartButtonProject.SmartButton sbLP_Close 
      Height          =   225
      Left            =   10995
      TabIndex        =   63
      ToolTipText     =   "Close Live Preview"
      Top             =   8985
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   397
      Picture         =   "frmMain.frx":2CFA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OffsetLeft      =   2
   End
   Begin SmartButtonProject.SmartButton sbLP_GroupMode 
      Height          =   225
      Left            =   10530
      TabIndex        =   62
      Tag             =   "TBI"
      ToolTipText     =   "Click to toggle between Toolbar Item preview and Group preview"
      Top             =   8985
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   397
      Picture         =   "frmMain.frx":3094
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OffsetRight     =   1
      OffsetTop       =   5
   End
   Begin VB.PictureBox picSplit2 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   120
      Left            =   1560
      MouseIcon       =   "frmMain.frx":362E
      MousePointer    =   7  'Size N S
      ScaleHeight     =   120
      ScaleWidth      =   2505
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   3945
      Width           =   2500
   End
   Begin SHDocVwCtl.WebBrowser wbMainPreview 
      CausesValidation=   0   'False
      Height          =   585
      Left            =   9000
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   8850
      Width           =   1365
      ExtentX         =   2408
      ExtentY         =   1032
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Timer tmrDelayedInitLivePreview 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   5175
      Top             =   9345
   End
   Begin VB.Timer tmrDelayedLivePreview 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4620
      Top             =   9345
   End
   Begin VB.PictureBox picEasy 
      Height          =   3870
      Left            =   8820
      ScaleHeight     =   3810
      ScaleWidth      =   5055
      TabIndex        =   45
      Top             =   4815
      Visible         =   0   'False
      Width           =   5115
      Begin VB.Frame frmEasySubMenu 
         Caption         =   "SubMenu Configuration"
         Height          =   960
         Left            =   0
         TabIndex        =   55
         Top             =   1545
         Width           =   4200
         Begin VB.ComboBox cmbEasySubMenuAction 
            Height          =   315
            ItemData        =   "frmMain.frx":3780
            Left            =   120
            List            =   "frmMain.frx":378A
            Style           =   2  'Dropdown List
            TabIndex        =   56
            Top             =   450
            Width           =   2100
         End
         Begin MSComctlLib.ImageCombo icmbEasyAlignment 
            Height          =   330
            Left            =   2460
            TabIndex        =   57
            Top             =   450
            Width           =   1500
            _ExtentX        =   2646
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
            ImageList       =   "ilAlignment"
         End
         Begin VB.Label Label1 
            BackStyle       =   0  'Transparent
            Caption         =   "Action"
            Height          =   195
            Left            =   135
            TabIndex        =   59
            Top             =   240
            Width           =   450
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alignment"
            Height          =   195
            Left            =   2475
            TabIndex        =   58
            Top             =   240
            Width           =   705
         End
      End
      Begin VB.Frame frmEasyLink 
         Caption         =   "Menu Link"
         Height          =   1440
         Left            =   0
         TabIndex        =   46
         Top             =   30
         Width           =   4200
         Begin VB.PictureBox picEasyLinkBtns 
            Height          =   405
            Left            =   135
            ScaleHeight     =   345
            ScaleWidth      =   3855
            TabIndex        =   50
            Top             =   915
            Width           =   3915
            Begin VB.CheckBox chkEasyEnableNewWindow 
               Caption         =   "Check1"
               Height          =   210
               Left            =   3540
               TabIndex        =   54
               Top             =   60
               Width           =   195
            End
            Begin SmartButtonProject.SmartButton sbEasyTargetFrame 
               Height          =   345
               Left            =   0
               TabIndex        =   51
               Top             =   0
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   609
               Caption         =   "       Frame"
               Picture         =   "frmMain.frx":37BD
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
            Begin SmartButtonProject.SmartButton sbEasyBookmark 
               Height          =   345
               Left            =   915
               TabIndex        =   52
               Top             =   0
               Width           =   1185
               _ExtentX        =   2090
               _ExtentY        =   609
               Caption         =   "       Bookmark"
               Picture         =   "frmMain.frx":3917
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
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
            Begin SmartButtonProject.SmartButton sbEasyNewWindow 
               Height          =   345
               Left            =   2115
               TabIndex        =   53
               Top             =   0
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   609
               Caption         =   "       New Window"
               Picture         =   "frmMain.frx":3EB1
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
         End
         Begin SmartButtonProject.SmartButton sbEasyLink 
            Height          =   315
            Left            =   3690
            TabIndex        =   49
            Top             =   465
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            Picture         =   "frmMain.frx":424B
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
         Begin VB.TextBox txtEasyLink 
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
            Left            =   135
            TabIndex        =   47
            ToolTipText     =   "Type the URL address of the link you want the browser to follow when this item is triggered"
            Top             =   450
            WhatsThisHelpID =   20070
            Width           =   3480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Link Address"
            Height          =   195
            Left            =   135
            TabIndex        =   48
            Top             =   240
            Width           =   900
         End
      End
   End
   Begin VB.Timer tmrDoReg 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   135
      Top             =   9345
   End
   Begin VB.CheckBox chkCompile 
      Caption         =   "Compile"
      Height          =   225
      Left            =   4575
      TabIndex        =   44
      ToolTipText     =   "Check this option to enable events for this command"
      Top             =   2010
      Width           =   1050
   End
   Begin VB.TextBox txtCaption 
      Height          =   300
      Left            =   3315
      TabIndex        =   43
      Top             =   1620
      Width           =   3135
   End
   Begin VB.Timer tmrResize 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   9015
      Top             =   7545
   End
   Begin VB.Timer tmrDelayCaptionUpdate 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   6300
      Top             =   9345
   End
   Begin VB.PictureBox picObj1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   6960
      ScaleHeight     =   330
      ScaleWidth      =   405
      TabIndex        =   42
      Top             =   7710
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   3315
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   5355
      Width           =   3135
   End
   Begin VB.Timer tmrInitStyleDlg 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4050
      Top             =   9345
   End
   Begin VB.Timer tmrDblCheck 
      Enabled         =   0   'False
      Interval        =   63000
      Left            =   1260
      Top             =   9345
   End
   Begin VB.Timer tmrDEMOInfo 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   690
      Top             =   9345
   End
   Begin VB.Frame frameEvent 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4635
      Left            =   7050
      TabIndex        =   20
      Top             =   1605
      Width           =   4350
      Begin SmartButtonProject.SmartButton cmdFindTargetGroup 
         Height          =   315
         Left            =   3195
         TabIndex        =   27
         Top             =   1350
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmMain.frx":43A5
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
      Begin VB.ComboBox cmbTargetMenu 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2130
         Width           =   2970
      End
      Begin SmartButtonProject.SmartButton cmdTargetFrame 
         Height          =   315
         Left            =   1965
         TabIndex        =   24
         Top             =   1140
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmMain.frx":44FF
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SmartButtonProject.SmartButton cmdBookmark 
         Height          =   315
         Left            =   2415
         TabIndex        =   25
         Top             =   1155
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmMain.frx":4659
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ImageCombo icmbAlignment 
         Height          =   330
         Left            =   405
         TabIndex        =   33
         Top             =   3930
         Width           =   3135
         _ExtentX        =   5530
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
         ImageList       =   "ilAlignment"
      End
      Begin VB.ComboBox cmbActionType 
         Height          =   315
         ItemData        =   "frmMain.frx":4BF3
         Left            =   180
         List            =   "frmMain.frx":4C03
         Style           =   2  'Dropdown List
         TabIndex        =   22
         ToolTipText     =   "Select the type of action to perform when the selected event is triggered"
         Top             =   615
         Width           =   3405
      End
      Begin VB.TextBox txtURL 
         Enabled         =   0   'False
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
         Left            =   165
         TabIndex        =   26
         ToolTipText     =   "Type the URL address of the link you want the browser to follow when this item is triggered"
         Top             =   1350
         WhatsThisHelpID =   20070
         Width           =   2970
      End
      Begin VB.ComboBox cmbTargetFrame 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmMain.frx":4C53
         Left            =   1395
         List            =   "frmMain.frx":4C55
         TabIndex        =   31
         ToolTipText     =   "Select the target frame for the action"
         Top             =   3330
         Visible         =   0   'False
         Width           =   2115
      End
      Begin SmartButtonProject.SmartButton cmdWinParams 
         Height          =   315
         Left            =   1470
         TabIndex        =   30
         Top             =   2910
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   556
         Caption         =   "Parameters"
         Picture         =   "frmMain.frx":4C57
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   5
         PictureLayout   =   3
         OffsetRight     =   5
      End
      Begin SmartButtonProject.SmartButton cmdBrowse 
         Height          =   315
         Left            =   3210
         TabIndex        =   29
         Top             =   2130
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmMain.frx":4FF1
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
      Begin VB.Label lblAlignment 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Alignment"
         Height          =   195
         Left            =   405
         TabIndex        =   32
         Top             =   3705
         Width           =   705
      End
      Begin VB.Label lblActionType 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Action Type"
         Height          =   195
         Left            =   210
         TabIndex        =   21
         Top             =   375
         Width           =   855
      End
      Begin VB.Label lblActionName 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Target Menu"
         Height          =   195
         Left            =   165
         TabIndex        =   23
         Top             =   1110
         Width           =   915
      End
   End
   Begin VB.CheckBox chkEnabled 
      Caption         =   "Enabled"
      Height          =   225
      Left            =   3315
      TabIndex        =   9
      ToolTipText     =   "Check this option to enable events for this command"
      Top             =   2010
      Width           =   1050
   End
   Begin VB.OptionButton opAlignmentStyle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3405
      Picture         =   "frmMain.frx":514B
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6525
      Width           =   315
   End
   Begin VB.OptionButton opAlignmentStyle 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3405
      Picture         =   "frmMain.frx":5295
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   6225
      Value           =   -1  'True
      Width           =   315
   End
   Begin VB.PictureBox picSplit 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2880
      Left            =   2535
      MouseIcon       =   "frmMain.frx":53DF
      MousePointer    =   9  'Size W E
      ScaleHeight     =   2880
      ScaleWidth      =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1230
      Width           =   120
   End
   Begin MSComctlLib.TreeView tvMenus 
      Height          =   4155
      Left            =   210
      TabIndex        =   8
      Top             =   1545
      WhatsThisHelpID =   20000
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   7329
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   1
      ImageList       =   "ilIcons"
      Appearance      =   0
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin MSComctlLib.TreeView tvMapView 
      Height          =   1230
      Left            =   330
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   7770
      WhatsThisHelpID =   20000
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   2170
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   617
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ilIcons"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   1815
      Top             =   9345
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6300
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   7740
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picItemIcon2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6555
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   36
      TabStop         =   0   'False
      Top             =   8130
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picItemIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   6930
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   8130
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picRsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   6330
      ScaleHeight     =   435
      ScaleWidth      =   495
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   7635
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Timer tmrLauncher 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3495
      Top             =   9345
   End
   Begin MSComctlLib.ImageList ilAlignment 
      Left            =   5010
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5531
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5ACB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6065
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":65FF
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6B99
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7133
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":76CD
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7C67
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8201
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":879B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D35
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":92CF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilTabs 
      Left            =   5025
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483633
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   6
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9869
            Key             =   "MoveON"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":99C3
            Key             =   "DblClickON"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9B1D
            Key             =   "ClickON"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9C77
            Key             =   "MoveOFF"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9D59
            Key             =   "DblClickOFF"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9E3B
            Key             =   "ClickOFF"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   4410
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   142
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9F1D
            Key             =   "mnuMenuColor|mnuContextColor"
            Object.Tag             =   "Color"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A079
            Key             =   "mnuMenuFont|mnuContextFont"
            Object.Tag             =   "Font"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A1D5
            Key             =   "mnuMenuCursor|mnuContextCursor"
            Object.Tag             =   "Cursor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A331
            Key             =   "mnuMenuImage|mnuContextImage"
            Object.Tag             =   "Image"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A48D
            Key             =   "mnuMenuSFX|mnuContextSFX"
            Object.Tag             =   "Special Effects"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A5E7
            Key             =   "btnUp"
            Object.Tag             =   "Up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A743
            Key             =   "btnDown"
            Object.Tag             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A89F
            Key             =   "mnuFileNew"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AE39
            Key             =   "mnuFileOpen"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AF93
            Key             =   "mnuFileSave"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B52D
            Key             =   "mnuEditCopy|mnuContextCopy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BAC7
            Key             =   "mnuEditPaste|mnuContextPaste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C061
            Key             =   "mnuEditUndo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C5FB
            Key             =   "mnuEditRedo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CB95
            Key             =   "mnuEditDelete|mnuContextDelete|mnuTBContextDelete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D12F
            Key             =   "mnuMenuAddGroup|mnuTBContextAddGroup"
            Object.Tag             =   "New Group"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D6CB
            Key             =   "mnuMenuAddCommand|mnuContextAddCommand"
            Object.Tag             =   "New Command"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DC67
            Key             =   "mnuMenuAddSeparator|mnuContextAddSeparator"
            Object.Tag             =   "New Separator"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E203
            Key             =   ""
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E79F
            Key             =   ""
            Object.Tag             =   "Command"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ED3B
            Key             =   ""
            Object.Tag             =   "Separator"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F2D7
            Key             =   "mnuToolsPreview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F873
            Key             =   "mnuToolsHotSpotsEditor"
            Object.Tag             =   "HotSpots Editor"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FE0F
            Key             =   "mnuFileProjProp"
            Object.Tag             =   "Properties"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":101A9
            Key             =   "mnuToolsCompile"
            Object.Tag             =   "Compile"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10FFB
            Key             =   "mnuHelpUpgrade"
            Object.Tag             =   "Upgrade"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115A3
            Key             =   "mnuHelpXFXFAQ"
            Object.Tag             =   "FAQ"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":118F7
            Key             =   "mnuHelpXFXNews"
            Object.Tag             =   "News"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11C4B
            Key             =   "mnuHelpXFXSupport"
            Object.Tag             =   "Support"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F9F
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":122F3
            Key             =   "mnuHelpXFXPublicForum"
            Object.Tag             =   "Forum"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12647
            Key             =   "mnuEditFind"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":127A1
            Key             =   "NoEvents"
            Object.Tag             =   "NoEvents"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12D3B
            Key             =   "OverCascade"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":132D5
            Key             =   "Click"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1386F
            Key             =   "ClickCascade"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13E09
            Key             =   "DoubleClick"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":143A3
            Key             =   "DoubleClickCascade"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1493D
            Key             =   "Over"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14ED7
            Key             =   "Disabled"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15471
            Key             =   "mnuMenuMargins|mnuContextMargins"
            Object.Tag             =   "Margins"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1580B
            Key             =   "mnuToolsPublish"
            Object.Tag             =   "Publish"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17FBD
            Key             =   "mnuMenuSound"
            Object.Tag             =   "Sounds"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18117
            Key             =   "GClick"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":186B1
            Key             =   "GOver"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18C4B
            Key             =   "GOverCascade"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":191E5
            Key             =   "GNoEvents"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1977F
            Key             =   "GClickCascade"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19D19
            Key             =   "GDoubleClick"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A2B3
            Key             =   "GDoubleClickCascade"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A84D
            Key             =   "GDisabled"
            Object.Tag             =   "DisabledGroup"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ADE7
            Key             =   "mnuEditRename|mnuContextRename|mnuTBContextRename"
            Object.Tag             =   "Rename"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B181
            Key             =   "mnuEditPreferences"
            Object.Tag             =   "Preferences"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B51B
            Key             =   "mnuFileRF"
            Object.Tag             =   "DMBProjectIcon"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E225
            Key             =   "EmptyIcon"
            Object.Tag             =   "EmptyIcon"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E5BF
            Key             =   "mnuToolsAddInEditor"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E959
            Key             =   ""
            Object.Tag             =   "Left"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EAB3
            Key             =   ""
            Object.Tag             =   "Right"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EC0D
            Key             =   ""
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ED67
            Key             =   ""
            Object.Tag             =   "Over"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EEC1
            Key             =   ""
            Object.Tag             =   "Font Bold"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EFD3
            Key             =   ""
            Object.Tag             =   "Font Italic"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F0E5
            Key             =   ""
            Object.Tag             =   "Font Underline"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F1F7
            Key             =   ""
            Object.Tag             =   "Size"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F309
            Key             =   ""
            Object.Tag             =   "Font Name"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F41B
            Key             =   ""
            Object.Tag             =   "Toolbar Item"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F575
            Key             =   ""
            Object.Tag             =   "Target Frame"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F6CF
            Key             =   ""
            Object.Tag             =   "Leading"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F829
            Key             =   ""
            Object.Tag             =   "Group Alignment"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FDC3
            Key             =   ""
            Object.Tag             =   "Caption Alignment"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FED5
            Key             =   ""
            Object.Tag             =   "Events"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2002F
            Key             =   ""
            Object.Tag             =   "Frame"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20189
            Key             =   ""
            Object.Tag             =   "Line"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20723
            Key             =   "mnuMenuSelFX"
            Object.Tag             =   "Highlight Effects"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20CBD
            Key             =   ""
            Object.Tag             =   "Transparency"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20E17
            Key             =   ""
            Object.Tag             =   "Shadow"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20F71
            Key             =   ""
            Object.Tag             =   "All Properties"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":210CB
            Key             =   ""
            Object.Tag             =   "EventOver"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21225
            Key             =   ""
            Object.Tag             =   "EventDoubleClick"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2137F
            Key             =   ""
            Object.Tag             =   "EventClick"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":214D9
            Key             =   ""
            Object.Tag             =   "URL"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21A73
            Key             =   ""
            Object.Tag             =   "Action Type"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2200D
            Key             =   ""
            Object.Tag             =   "Border Size"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":223A7
            Key             =   ""
            Object.Tag             =   "Command Horizontal Margin"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22741
            Key             =   ""
            Object.Tag             =   "Command Vertical Margin"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22ADB
            Key             =   ""
            Object.Tag             =   "Vertical Margin"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22E75
            Key             =   ""
            Object.Tag             =   "Horizontal Margin"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2320F
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":235A9
            Key             =   ""
            Object.Tag             =   "New Window"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23943
            Key             =   ""
            Object.Tag             =   "Colored Borders"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23CDD
            Key             =   ""
            Object.Tag             =   "Group Width"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24077
            Key             =   ""
            Object.Tag             =   "Group Height"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24411
            Key             =   ""
            Object.Tag             =   "Text Color"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":249AB
            Key             =   ""
            Object.Tag             =   "Back Color"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24F45
            Key             =   ""
            Object.Tag             =   "Status Text"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":254DF
            Key             =   ""
            Object.Tag             =   "Caption"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":25A79
            Key             =   "mnuRegisterUnlock"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26753
            Key             =   ""
            Object.Tag             =   "Commands Layout"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26AED
            Key             =   ""
            Object.Tag             =   "Group Effects"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27087
            Key             =   ""
            Object.Tag             =   "Overlay"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27621
            Key             =   ""
            Object.Tag             =   "HS-Text"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":27BBB
            Key             =   ""
            Object.Tag             =   "HS-Image"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28155
            Key             =   ""
            Object.Tag             =   "HS-DynaText"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":286EF
            Key             =   "mnuHelpContents"
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":28C89
            Key             =   "www"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29223
            Key             =   "mnuMenuToolbarProperties|mnuToolsToolbarsEditor|mnuTBContextToolbarProperties"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2937D
            Key             =   "mnuToolsLCMan|mnuToolsLC"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":294D7
            Key             =   "mnuMenuAddSubGroup|mnuContextAddSubGroup"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29A71
            Key             =   "mnuMenuLength|mnuContextLength"
            Object.Tag             =   "SeparatorLength"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A00B
            Key             =   ""
            Object.Tag             =   "Justify"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2A5A5
            Key             =   ""
            Object.Tag             =   "Toolbar Alignment"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AB3F
            Key             =   ""
            Object.Tag             =   "Spanning"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AC99
            Key             =   ""
            Object.Tag             =   "Follow Scrolling"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2ADF3
            Key             =   ""
            Object.Tag             =   "Toolbar Offset"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2AF4D
            Key             =   "mnuMenuAddToolbar|mnuTBContextAddToolbar"
            Object.Tag             =   "Add Toolbar"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B0A7
            Key             =   "mnuMenuRemoveToolbar"
            Object.Tag             =   "Remove Toolbar"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B201
            Key             =   "mnuEditFindReplace"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2B79B
            Key             =   "mnuHelpSearch"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2BD35
            Key             =   "mnuHelpXFXHomePage"
            Object.Tag             =   "xFXLogo"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C0CF
            Key             =   ""
            Object.Tag             =   "Scrolling"
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C469
            Key             =   "mnuFileNewEmpty"
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CA03
            Key             =   "mnuFileNewFromDir"
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2CF9D
            Key             =   "mnuFileNewFromPreset"
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2D537
            Key             =   "mnuFileNewFromWizard"
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DAD1
            Key             =   "Transparent"
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2DC2B
            Key             =   ""
            Object.Tag             =   "Right Align"
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E1C5
            Key             =   ""
            Object.Tag             =   "Center Align"
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E75F
            Key             =   ""
            Object.Tag             =   "Left Align"
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2ECF9
            Key             =   "TDoubleClickCascade"
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EE53
            Key             =   "TNoEvents"
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2EFAD
            Key             =   "TOver"
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F107
            Key             =   "TOverCascade"
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F261
            Key             =   "TClick"
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F3BB
            Key             =   "TClickCascade"
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F515
            Key             =   "TDoubleClick"
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2F66F
            Key             =   "mnuToolsIHW"
            Object.Tag             =   "Visibility Condition"
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FC09
            Key             =   "mnuToolsReports"
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":301A3
            Key             =   "mnuToolsInstallMenus|mnuToolsInstallMenusA"
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":302FD
            Key             =   "mnuToolsApplyStyle"
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30697
            Key             =   "HelpF1"
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":30C31
            Key             =   "mnuToolsBrokenLinks"
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":311CB
            Key             =   "mnuHelpDLPDF"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1155
      ScaleHeight     =   255
      ScaleWidth      =   3270
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   8655
      Visible         =   0   'False
      Width           =   3270
   End
   Begin VB.TextBox dummyText 
      Height          =   315
      Left            =   2055
      MultiLine       =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   8775
      Visible         =   0   'False
      Width           =   3570
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   5670
      Top             =   7500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "0"
      DialogTitle     =   "Select Project"
      Filter          =   "DHTML Menu Builder Projects|*.dmb"
      MaxFileSize     =   1024
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   2940
      Top             =   9345
   End
   Begin MSComctlLib.StatusBar sbDummy 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   41
      Top             =   9840
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   19994
            MinWidth        =   4410
            Text            =   "Sel Info"
            TextSave        =   "Sel Info"
            Key             =   "sbFlood"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   4419
            MinWidth        =   4410
            Text            =   "Config Info"
            TextSave        =   "Config Info"
            Key             =   "sbConfig"
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
      OLEDropMode     =   1
   End
   Begin MSComctlLib.Toolbar tbCmd 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   1
      Top             =   360
      WhatsThisHelpID =   20100
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbColor"
            Object.ToolTipText     =   "Color"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbFont"
            Object.ToolTipText     =   "Font"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbCursor"
            Object.ToolTipText     =   "Cursor"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbImage"
            Object.ToolTipText     =   "Image"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   730
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbMargins"
            Object.ToolTipText     =   "Margins"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbSelFX"
            Object.ToolTipText     =   "Selection Effects"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbFX"
            Object.ToolTipText     =   "Special Effects"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbSound"
            Object.ToolTipText     =   "Sound"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbUp"
            Object.ToolTipText     =   "Move selected item Up"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbDown"
            Object.ToolTipText     =   "Move selected item Down"
         EndProperty
      EndProperty
      Begin VB.PictureBox picLeading 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1425
         ScaleHeight     =   330
         ScaleWidth      =   675
         TabIndex        =   2
         ToolTipText     =   "Leading"
         Top             =   0
         WhatsThisHelpID =   20030
         Width           =   675
         Begin MSComCtl2.UpDown udLeading 
            Height          =   240
            Left            =   375
            TabIndex        =   3
            ToolTipText     =   "Leading"
            Top             =   45
            WhatsThisHelpID =   20030
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   423
            _Version        =   393216
            Value           =   3
            Max             =   5
            Enabled         =   0   'False
         End
         Begin VB.Shape shpSpc 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   60
            Index           =   2
            Left            =   60
            Top             =   210
            Width           =   255
         End
         Begin VB.Shape shpSpc 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   60
            Index           =   1
            Left            =   60
            Top             =   90
            Width           =   255
         End
         Begin VB.Shape shpSpc 
            BorderStyle     =   0  'Transparent
            FillColor       =   &H00808080&
            FillStyle       =   0  'Solid
            Height          =   60
            Index           =   0
            Left            =   60
            Top             =   150
            Width           =   255
         End
         Begin VB.Line ln1 
            BorderColor     =   &H00FFFFFF&
            Index           =   0
            Visible         =   0   'False
            X1              =   0
            X2              =   895
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line ln1 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            Visible         =   0   'False
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   360
         End
         Begin VB.Line ln1 
            BorderColor     =   &H00808080&
            Index           =   3
            Visible         =   0   'False
            X1              =   -165
            X2              =   925
            Y1              =   315
            Y2              =   315
         End
         Begin VB.Line ln1 
            BorderColor     =   &H00808080&
            Index           =   2
            Visible         =   0   'False
            X1              =   660
            X2              =   660
            Y1              =   0
            Y2              =   360
         End
      End
   End
   Begin MSComctlLib.Toolbar tbMenu2 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   4
      Top             =   720
      WhatsThisHelpID =   20110
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbCompile"
            Object.ToolTipText     =   "Compile"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbPublish"
            Object.ToolTipText     =   "Publish"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbSep01"
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbPreview"
            Object.ToolTipText     =   "Preview"
            Style           =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbHotSpotEditor"
            Object.ToolTipText     =   "HotSpot Editor"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbTBEditor"
            Object.ToolTipText     =   "Toolbars Editor"
            Style           =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbProperties"
            Object.ToolTipText     =   "Project Properties"
            Style           =   5
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      WhatsThisHelpID =   20115
      Width           =   14160
      _ExtentX        =   24977
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   19
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbNew"
            Object.ToolTipText     =   "New Project"
            Style           =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbOpen"
            Object.ToolTipText     =   "Open Project"
            Style           =   5
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbSave"
            Object.ToolTipText     =   "Save Project"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbAddGrp"
            Object.ToolTipText     =   "Add Group"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbAddSubGrp"
            Object.ToolTipText     =   "Add SubGroup"
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbAddCmd"
            Object.ToolTipText     =   "Add Command"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbAddSep"
            Object.ToolTipText     =   "Add Separator"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbCopy"
            Object.ToolTipText     =   "Copy"
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbPaste"
            Object.ToolTipText     =   "Paste"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbFind"
            Object.ToolTipText     =   "Find"
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbUndo"
            Style           =   5
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbRedo"
            Style           =   5
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbRemove"
            Object.ToolTipText     =   "Delete"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TabStrip tsCmdType 
      Height          =   2655
      Left            =   3300
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2400
      WhatsThisHelpID =   20050
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   4683
      TabWidthStyle   =   1
      ShowTips        =   0   'False
      Separators      =   -1  'True
      TabMinWidth     =   706
      ImageList       =   "ilTabs"
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Over"
            Key             =   "tsOver"
            ImageVarType    =   2
            ImageIndex      =   1
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Click"
            Key             =   "tsClick"
            ImageVarType    =   2
            ImageIndex      =   3
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Double Click"
            Key             =   "tsDoubleClick"
            ImageVarType    =   2
            ImageIndex      =   2
         EndProperty
      EndProperty
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
   End
   Begin VLMnuPlus.VLMenuPlus vlmCtrl 
      Left            =   6075
      Top             =   6300
      _ExtentX        =   847
      _ExtentY        =   847
      _CXY            =   4
      _CGUID          =   43495.5104166667
      BitmapBackground=   -2147483644
      Language        =   0
   End
   Begin VB.Label lblTSViewsNormal 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Normal"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   660
      TabIndex        =   15
      Top             =   6195
      Width           =   780
   End
   Begin VB.Label lblTSViewsMap 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Map"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   735
      TabIndex        =   13
      Top             =   5985
      Width           =   780
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   3315
      TabIndex        =   7
      Top             =   1380
      Width           =   555
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status Text"
      Height          =   195
      Left            =   3315
      TabIndex        =   11
      Top             =   5130
      Width           =   840
   End
   Begin VB.Label lblASHorizontal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horizontal"
      Height          =   195
      Left            =   3825
      TabIndex        =   19
      Top             =   6555
      Width           =   720
   End
   Begin VB.Label lblASVertical 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vertical"
      Height          =   195
      Left            =   3825
      TabIndex        =   17
      Top             =   6255
      Width           =   525
   End
   Begin VB.Label lblLayout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Commands Layout"
      Height          =   195
      Left            =   3315
      TabIndex        =   14
      Top             =   5985
      Width           =   1320
   End
   Begin VB.Label lblDataTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group Properties"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   3555
      TabIndex        =   5
      Top             =   1050
      Width           =   1440
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Begin VB.Menu mnuFileNewEmpty 
            Caption         =   "Empty Project"
         End
         Begin VB.Menu mnuFileNewSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileNewFromPreset 
            Caption         =   "From Preset..."
         End
         Begin VB.Menu mnuFileNewFromWizard 
            Caption         =   "Using the Wizard..."
         End
         Begin VB.Menu mnuFileNewFromDir 
            Caption         =   "From Directory Structure..."
         End
         Begin VB.Menu mnuFileNewFromROR 
            Caption         =   "Import from ROR File..."
         End
         Begin VB.Menu mnuFileNewFromTXT 
            Caption         =   "Import from Text File..."
         End
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpenRecent 
         Caption         =   "Open Recent"
         Begin VB.Menu mnuFileOpenRecentR 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu mnuFileOpenRecentOPSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuFileOpenRecentOP 
            Caption         =   "Open Last Saved Project"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuFileSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSaveAs 
         Caption         =   "Save As..."
      End
      Begin VB.Menu mnuFileSaveAsPreset 
         Caption         =   "Save As Preset"
      End
      Begin VB.Menu mnuFileSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSubmitPreset 
         Caption         =   "Submit Preset..."
      End
      Begin VB.Menu mnuFileSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExportHTML 
         Caption         =   "Export As HTML..."
      End
      Begin VB.Menu mnuFileExportSitemap 
         Caption         =   "Export As Sitemap..."
      End
      Begin VB.Menu mnuFileSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProjProp 
         Caption         =   "Project Properties..."
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileSep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "Edit"
      Begin VB.Menu mnuEditUndo 
         Caption         =   "Undo"
         Shortcut        =   ^Z
      End
      Begin VB.Menu mnuEditRedo 
         Caption         =   "Redo"
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuEditSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditCopy 
         Caption         =   "Copy..."
         Shortcut        =   ^{INSERT}
      End
      Begin VB.Menu mnuEditPaste 
         Caption         =   "Paste..."
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuEditSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditFind 
         Caption         =   "Find..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuEditFindNext 
         Caption         =   "Find Next"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuEditFindReplace 
         Caption         =   "Replace..."
      End
      Begin VB.Menu mnuEditSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuEditRename 
         Caption         =   "Rename"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuEditSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPreferences 
         Caption         =   "Preferences..."
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Begin VB.Menu mnuMenuAddToolbar 
         Caption         =   "Add Toolbar"
      End
      Begin VB.Menu mnuMenuRemoveToolbar 
         Caption         =   "Remove Toolbar"
      End
      Begin VB.Menu mnuMenuSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuAddGroup 
         Caption         =   "Add Group"
         Shortcut        =   {F6}
      End
      Begin VB.Menu mnuMenuAddSubGroup 
         Caption         =   "Add SubGroup"
      End
      Begin VB.Menu mnuMenuSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuAddCommand 
         Caption         =   "Add Command"
         Shortcut        =   {F7}
      End
      Begin VB.Menu mnuMenuAddSeparator 
         Caption         =   "Add Separator"
         Shortcut        =   +{F7}
      End
      Begin VB.Menu mnuMenuSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuCopy 
         Caption         =   "Copy..."
      End
      Begin VB.Menu mnuMenuPaste 
         Caption         =   "Paste..."
      End
      Begin VB.Menu mnuMenuSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuDelete 
         Caption         =   "Delete"
      End
      Begin VB.Menu mnuMenuRename 
         Caption         =   "Rename"
      End
      Begin VB.Menu mnuMenuSep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuColor 
         Caption         =   "Color..."
         Shortcut        =   ^{F1}
      End
      Begin VB.Menu mnuMenuFont 
         Caption         =   "Font..."
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu mnuMenuCursor 
         Caption         =   "Cursor..."
         Shortcut        =   ^{F3}
      End
      Begin VB.Menu mnuMenuImage 
         Caption         =   "Image..."
         Shortcut        =   ^{F4}
      End
      Begin VB.Menu mnuMenuMargins 
         Caption         =   "Margins"
         Shortcut        =   ^{F5}
      End
      Begin VB.Menu mnuMenuSelFX 
         Caption         =   "Selection Effects"
         Shortcut        =   ^{F6}
      End
      Begin VB.Menu mnuMenuSFX 
         Caption         =   "Special Effects"
         Shortcut        =   ^{F7}
      End
      Begin VB.Menu mnuMenuLength 
         Caption         =   "Length..."
      End
      Begin VB.Menu mnuMenuSep06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMenuToolbarProperties 
         Caption         =   "Toolbar Properties"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsPreview 
         Caption         =   "Preview..."
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuToolsLivePreview 
         Caption         =   "Live Preview"
      End
      Begin VB.Menu mnuToolsSetDefaultBrowser 
         Caption         =   "Set Default Browser..."
      End
      Begin VB.Menu mnuToolsSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsCompile 
         Caption         =   "Compile..."
      End
      Begin VB.Menu mnuToolsPublish 
         Caption         =   "Publish..."
      End
      Begin VB.Menu mnuToolsSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsInstallMenus 
         Caption         =   "Install Menus..."
      End
      Begin VB.Menu mnuToolsInstallMenusA 
         Caption         =   "Install Menus"
         Begin VB.Menu mnuToolsInstallMenusAILC 
            Caption         =   "Install Loader Code"
         End
         Begin VB.Menu mnuToolsInstallMenusAIFLC 
            Caption         =   "Install Frames Loader Code"
         End
         Begin VB.Menu mnuToolsInstallMenusSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuToolsInstallMenusAIRLC 
            Caption         =   "Remove Loader Code"
         End
      End
      Begin VB.Menu mnuToolsSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsDefaultConfig 
         Caption         =   "Default Configuration..."
      End
      Begin VB.Menu mnuToolsSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsHotSpotsEditor 
         Caption         =   "HotSpots Editor..."
      End
      Begin VB.Menu mnuToolsToolbarsEditor 
         Caption         =   "Toolbars Editor..."
      End
      Begin VB.Menu mnuToolsAddInEditor 
         Caption         =   "AddIn Editor..."
      End
      Begin VB.Menu mnuToolsSep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsDynAPI 
         Caption         =   "DynAPI"
      End
      Begin VB.Menu mnuToolsSep06 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsSecProj 
         Caption         =   "Secondary Projects..."
      End
      Begin VB.Menu mnuToolsApplyStyle 
         Caption         =   "Apply Style from Preset..."
      End
      Begin VB.Menu mnuToolsIHW 
         Caption         =   "Item Highlight Wizard..."
      End
      Begin VB.Menu mnuToolsExtractIcon 
         Caption         =   "Extract Icon..."
      End
      Begin VB.Menu mnuToolsSep09 
         Caption         =   "-"
      End
      Begin VB.Menu mnuToolsReports 
         Caption         =   "Reports"
         Begin VB.Menu mnuToolsReport 
            Caption         =   "Compilation Report..."
         End
         Begin VB.Menu mnuToolsBrokenLinks 
            Caption         =   "Broken Links Report..."
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents..."
         Shortcut        =   +{F1}
      End
      Begin VB.Menu mnuHelpTutorials 
         Caption         =   "Tutorials"
         Begin VB.Menu mnuHelpTutotialsTB 
            Caption         =   "Using Toolbars..."
         End
         Begin VB.Menu mnuHelpPosMenus 
            Caption         =   "Positioning the menus..."
         End
         Begin VB.Menu mnuHelpTutorialsHS 
            Caption         =   "Using HotSpots..."
         End
         Begin VB.Menu mnuHelpTutorialsSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelpTutorialsVideo 
            Caption         =   "Video Tutorial..."
         End
      End
      Begin VB.Menu mnuHelpSearch 
         Caption         =   "Search..."
      End
      Begin VB.Menu mnuHelpSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpUpgrade 
         Caption         =   "Upgrade..."
      End
      Begin VB.Menu mnuHelpSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpDLPDF 
         Caption         =   "Download Documentation in PDF format..."
      End
      Begin VB.Menu mnuHelpSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpDynAPIDoc 
         Caption         =   "DynAPI Documentation..."
      End
      Begin VB.Menu mnuHelpSep05 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpXFX 
         Caption         =   "xFX JumpStart on the Web"
         Begin VB.Menu mnuHelpXFXHomePage 
            Caption         =   "Home Page"
         End
         Begin VB.Menu mnuHelpxFXWebSep01 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelpXFXSupport 
            Caption         =   "Support"
         End
         Begin VB.Menu mnuHelpXFXFAQ 
            Caption         =   "FAQ"
         End
         Begin VB.Menu mnuHelpXFXPublicForum 
            Caption         =   "Public Forum"
         End
         Begin VB.Menu mnuHelpxFXWeb02 
            Caption         =   "-"
         End
         Begin VB.Menu mnuHelpXFXNews 
            Caption         =   "News"
         End
      End
      Begin VB.Menu mnuHelpSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuRegister 
      Caption         =   "Register"
      Begin VB.Menu mnuRegisterUnlock 
         Caption         =   "Enter Registration Information"
      End
      Begin VB.Menu mnuRegisterBuy 
         Caption         =   "Purchase"
      End
   End
   Begin VB.Menu mnuBrowsers 
      Caption         =   "mnuBrowsers"
      Begin VB.Menu mnuBrowsersSetDefBrowser 
         Caption         =   "Set Default Browser..."
      End
      Begin VB.Menu mnuBrowsersSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBrowsersList 
         Caption         =   "Internal..."
         Index           =   0
      End
   End
   Begin VB.Menu mnuStyleOptions 
      Caption         =   "mnuStyleOptions"
      Begin VB.Menu mnuStyleOptionsOP 
         Caption         =   "Apply to select Command/Group"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuStyleOptionsOP 
         Caption         =   "Apply to all Commands under the GRP Group"
         Index           =   1
      End
      Begin VB.Menu mnuStyleOptionsOP 
         Caption         =   "Apply to all Commands/Groups under the TB Toolbar"
         Index           =   2
      End
      Begin VB.Menu mnuStyleOptionsOP 
         Caption         =   "Apply to all Commands/Groups in the Project"
         Index           =   3
      End
      Begin VB.Menu mnuStyleOptionsOPSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuStyleOptionsOPAdvanced 
         Caption         =   "Advanced..."
      End
   End
   Begin VB.Menu mnuPPShortcuts 
      Caption         =   "mnuPPShortcuts"
      Begin VB.Menu mnuPPShortcutsGeneral 
         Caption         =   "General"
      End
      Begin VB.Menu mnuPPShortcutsConfig 
         Caption         =   "Configurations"
      End
      Begin VB.Menu mnuPPShortcutsGlobal 
         Caption         =   "Global Settings"
      End
      Begin VB.Menu mnuPPShortcutsAdvanced 
         Caption         =   "Advanced"
      End
   End
   Begin VB.Menu mnuTBEShortcuts 
      Caption         =   "mnuTBEShortcuts"
      Begin VB.Menu mnuTBEShortcutsGeneral 
         Caption         =   "General"
      End
      Begin VB.Menu mnuTBEShortcutsAppearance 
         Caption         =   "Appearance"
      End
      Begin VB.Menu mnuTBEShortcutsPositioning 
         Caption         =   "Positioning"
      End
      Begin VB.Menu mnuTBEShortcutsEffects 
         Caption         =   "Effects"
      End
      Begin VB.Menu mnuTBEShortcutsAdvanced 
         Caption         =   "Advanced"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsUpdating As Boolean
Private IsLoadingProject As Boolean
Private DisableNodeClickEvent As Boolean
Private IsResizing As Boolean
Private IsRefreshingMap As Boolean
Private RenamingItem As String
Private InitDir As String

Private nodexExp() As String
Private IsRestoringExp As Boolean

Private Type RecentFile
    Title As String
    Path As String
End Type
Private RecentFiles() As RecentFile

Private SelToolbar As ToolBar

Private DragCmd As Integer
Private DragGrp As Integer

Private Enum SelectiveLauncherConstants
    [slcSelCopy]
    [slcSelPaste]
    [slcSelColor]
    [slcSelFont]
    [slcSelImage]
    [slcSelCursor]
    [slcSelMargin]
    [slcSelSFX]
    [slcSelSound]
    [slcSelSepLen]
    [slcSelCallMenu]
    [slcSelTBEditor]
    [slcShowContextMenu]
    [slcSelSelFX]
    [slcProjectProperties]
    [slcNone]
End Enum
Private SelectiveLauncher As SelectiveLauncherConstants

Private IsSplitting As Boolean
Private LastSelected As String
Private LastItemFP As String
Private IsRenaming As Boolean
Private IsSHIFT As Boolean

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1
Private WithEvents sbF1 As SmartButton
Attribute sbF1.VB_VarHelpID = -1
Private wbPreview As WebBrowser
Private LivePreviewIsBusy As Boolean

Private xMenu As CMenu

Private tvmvNode As Node

Private Declare Function GetUserDefaultLCID Lib "kernel32" () As Long

Private Sub chkCompile_Click()

    If IsUpdating Then Exit Sub
    Dim sNode As String
    
    #If LITE = 1 Then
        ShowLITELImitationInfo 2
    #Else
        If tvMapView.SelectedItem Is Nothing Then Exit Sub
        
        sNode = tvMapView.SelectedItem.Text
        
        UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(232)
        UpdateControls
        RefreshMap
        
        If InMapMode Then
            Dim nNode As Node
            For Each nNode In tvMapView.Nodes
                If nNode.Text = sNode Then
                    nNode.EnsureVisible
                    nNode.Selected = True
                    Exit For
                End If
            Next nNode
            
            sNode = tvMapView.SelectedItem.Text
            With tvMenus.SelectedItem
                UpdateStatusbar IsCommand(.key), IsGroup(.key), IsSeparator(.key)
            End With
        End If
    #End If
    
End Sub

Private Sub chkEasyEnableNewWindow_Click()

    Dim IsG As Boolean
    Dim IsC As Boolean
    Dim IsS As Boolean
    
    If tvMenus.SelectedItem Is Nothing Or IsTBMapSel Then
        Exit Sub
    Else
        IsG = IsGroup(tvMenus.SelectedItem.key)
        IsC = IsCommand(tvMenus.SelectedItem.key)
        IsS = IsSeparator(tvMenus.SelectedItem.key)
    End If
    
    If IsS Then Exit Sub
    
    If IsC Then
        With MenuCmds(GetID)
            .Actions.onclick.Type = IIf(chkEasyEnableNewWindow.Value = vbChecked, atcNewWindow, atcURL)
            If .Actions.onclick.Type = atcURL Then
                If .Actions.onclick.url = "" Then
                    .Actions.onclick.Type = atcNone
                End If
            End If
        End With
    Else
        With MenuGrps(GetID)
            .Actions.onclick.Type = IIf(chkEasyEnableNewWindow.Value = vbChecked, atcNewWindow, atcURL)
            If .Actions.onclick.Type = atcURL Then
                If .Actions.onclick.url = "" Then
                    .Actions.onclick.Type = atcNone
                End If
            End If
        End With
    End If
    
    UpdateControls

    'If chkEasyEnableNewWindow.Value = vbChecked Then sbEasyNewWindow_Click

End Sub

Private Sub chkEnabled_Click()

    If IsUpdating Then Exit Sub

    UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(232)
    UpdateControls
    RefreshMap

End Sub

Private Sub cmbActionType_Click()

    Dim lastAction As Integer
    
    If IsUpdating Then Exit Sub
    
    If IsGroup(tvMenus.SelectedItem.key) Then
        Select Case tsCmdType.SelectedItem.key
            Case "tsOver"
                lastAction = MenuGrps(GetID).Actions.onmouseover.Type
            Case "tsClick"
                lastAction = MenuGrps(GetID).Actions.onclick.Type
            Case "tsDoubleClick"
                lastAction = MenuGrps(GetID).Actions.OnDoubleClick.Type
        End Select
    Else
        Select Case tsCmdType.SelectedItem.key
            Case "tsOver"
                lastAction = MenuCmds(GetID).Actions.onmouseover.Type
            Case "tsClick"
                lastAction = MenuCmds(GetID).Actions.onclick.Type
            Case "tsDoubleClick"
                lastAction = MenuCmds(GetID).Actions.OnDoubleClick.Type
        End Select
    End If

    UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(108)
    UpdateControls
    ResizeURLControls
    
    If lastAction = 1 And cmbActionType.ListIndex = 3 Then Exit Sub
    If lastAction = 3 And cmbActionType.ListIndex = 1 Then Exit Sub
    If lastAction = cmbActionType.ListIndex Then Exit Sub
    RefreshMap
    
End Sub

Private Sub cmbTargetFrame_Change()

    If IsUpdating Then Exit Sub
    cmbTargetFrame_Click

End Sub

Private Sub cmbTargetFrame_Click()

    DontRefreshMap = True
    UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(235)

End Sub

Private Sub cmbTargetMenu_Click()

    If IsUpdating Then Exit Sub
    cmdFindTargetGroup.Enabled = (cmbTargetMenu.ListIndex >= 0)
    UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(109)
    RefreshMap

End Sub

Friend Sub ShowPreview()

    PreviewMode = pmcNormal

    If SomethingIsWrong(True, False) Or _
        InvalidItemsNames Or _
        DuplicatedItemsNames Or _
        InvalidItemsData Or _
        MissingParameters Then GoTo AbortPreview
    If Not MissingParameters Then
        If (Not DoCompile(True, , PreviewPath)) Then Exit Sub
        
        Err.Clear
        On Error Resume Next
        If IsWin9x Then
            frmPreview.Show vbModeless, Me
        Else
            frmPreview.Show vbModeless
        End If
        If Err.number = 401 Then
            Unload frmPreview
            frmPreview.Show vbModal
        Else
            frmPreview.tmrLoadPreview.Enabled = True
        End If
    End If
    
AbortPreview:

End Sub

Private Sub cmdBookmark_Click()

    frmURLBookmark.Show vbModal
    SetBookmarkState

End Sub

Friend Sub cmdBrowse_Click()

    Dim ActionName As String
    Dim FileName As String
    Dim sStr As String
    Dim ItemName As String
    
    If IsGroup(tvMenus.SelectedItem.key) Then
        ItemName = NiceGrpCaption(GetID)
    Else
        ItemName = NiceCmdCaption(GetID)
    End If
    ItemName = "'" + ItemName + "'"

    If LenB(InitDir) = 0 Or Left(InitDir, Len(GetRealLocal.RootWeb)) <> GetRealLocal.RootWeb Then
        InitDir = GetRealLocal.RootWeb
    End If
    With cDlg
        Select Case tsCmdType.SelectedItem.key
            Case "tsClick"
                ActionName = ItemName + " " + GetLocalizedStr(236)
                sStr = GetFilePath(Replace(txtURL.Text, Project.UserConfigs(Project.DefaultConfig).RootWeb, GetRealLocal.RootWeb))
                If FolderExists(sStr) And LenB(sStr) <> 0 Then
                    InitDir = sStr
                    FileName = GetFileName(txtURL.Text)
                End If
            Case "tsOver"
                ActionName = GetLocalizedStr(237) + " " + ItemName
            Case "tsDoubleClick"
                ActionName = tvMenus.SelectedItem.Text + " " + GetLocalizedStr(238)
        End Select
        .DialogTitle = GetLocalizedStr(239) + " " + ActionName
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNNoReadOnlyReturn
        .filter = SupportedHTMLDocs
        .InitDir = InitDir
        .FileName = FileName
        Err.Clear
        On Error Resume Next
        .ShowOpen
        If .FileName = FileName And LenB(FileName) <> 0 Then .FileName = InitDir + FileName
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then Exit Sub
        txtURL.Text = ConvertPath(.FileName)
        InitDir = GetFilePath(.FileName)
    End With

End Sub

Private Sub cmdFindTargetGroup_Click()

    On Error Resume Next
    
    Dim sText As String
    Dim i As Integer
    
    sText = cmbTargetMenu.Text
    If LenB(sText) = 0 Then Exit Sub
    
    If Left(sText, 1) = "[" Then
        sText = Mid(sText, 2, Len(sText) - 2)
    Else
        sText = MenuGrps(cmbTargetMenu.ListIndex + 1).Name
    End If
    
    LastItemFP = ""
    SelectItem tvMenus.Nodes("G" & GetIDByName(sText))
    If InMapMode Then
        SynchViews
        SetCtrlFocus tvMapView
    Else
        SetCtrlFocus tvMenus
    End If

End Sub

Private Sub cmdTargetFrame_Click()

    frmURLTargetFrame.Show vbModal
    SetBookmarkState

End Sub

Private Sub cmdWinParams_Click()

    If IsUpdating Then Exit Sub
    
    frmWindowOpenParams.Show vbModal
    
    UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(240)
    'UpdateControls

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 93 Then
        Select Case ActiveControl.Name
            Case tvMapView.Name
                Set tvmvNode = tvMapView.SelectedItem
                NodeSelectedInMapView tvmvNode
            Case ActiveControl.Name = tvMenus.Name
                Set tvmvNode = tvMenus.SelectedItem
            Case Else
                Exit Sub
        End Select
        If Not tvmvNode Is Nothing Then
            HandleContextMenu tvmvNode, vbRightButton
            Exit Sub
        End If
    End If

    If KeyCode = vbKeyF1 And Shift = 0 Then showHelp "dialogs/main.htm"
    
    If KeyCode = vbKeyDelete Then
        If ActiveControl.Name = "tvMapView" Or ActiveControl.Name = "tvMenus" Then mnuEditDelete_Click
    End If
    
    IsSHIFT = (Shift And vbShiftMask) = vbShiftMask

End Sub

Private Sub LoadPrgPrefs()
    
    If IsDEMO Then
        USER = "DEMO"
        COMPANY = "DEMO"
        USERSN = ""
    Else
        If GetSetting(App.EXEName, "RegInfo", "User", "DEMO") = "DEMO" Then
            If FileExists(AppPath + "rsc/reg.dat") Then
                On Error Resume Next
                
                Dim c As String
                c = UnCompress(AppPath + "rsc/reg.dat")
                SaveSetting App.EXEName, "RegInfo", "CacheData", Split(c, "|")(0)
                SaveSetting App.EXEName, "RegInfo", "CacheSig01", Split(c, "|")(1)
                SaveSetting App.EXEName, "RegInfo", "CacheSig02", Split(c, "|")(2)
                
                c = HEX2Str(Inflate(HEX2Str(Split(c, "|")(0))))
                SaveSetting App.EXEName, "RegInfo", "User", Split(c, "|")(0)
                SaveSetting App.EXEName, "RegInfo", "Company", Split(c, "|")(1)
                SaveSetting App.EXEName, "RegInfo", "SerialNumber", Split(c, "|")(2)
                SaveSetting App.EXEName, "RegInfo", "SubSystem", Split(c, "|")(4)
            End If
        End If
        
        USER = GetSetting(App.EXEName, "RegInfo", "User", "DEMO")
        COMPANY = GetSetting(App.EXEName, "RegInfo", "Company", "DEMO")
        USERSN = GetSetting(App.EXEName, "RegInfo", "SerialNumber", "")
        ORDERNUMBER = GetSetting(App.EXEName, "RegInfo", "OrderNumber", "")
    End If
    
    With Preferences
        .AutoRecover = CBool(GetSetting(App.EXEName, "Preferences", "AutoRecover", 1))
        .OpenLastProject = CBool(GetSetting(App.EXEName, "Preferences", "OpenLastProject", 1))
        .SepHeight = GetSetting(App.EXEName, "Preferences", "SepHeight", 11)
        If IsDEMO Then
            .ShowNag = True
        Else
            .ShowNag = CBool(GetSetting(App.EXEName, "Preferences", "ShowNag", 1))
        End If
        .ShowWarningAddInEditor = CBool(GetSetting(App.EXEName, "Preferences", "ShowWarningAIE", 1))
        .ShowPPOnNewProject = CBool(GetSetting(App.EXEName, "Preferences", "ShowPPOnNewProject", 1))
        
        .CommandsInheritance = GetSetting(App.EXEName, "Preferences", "CmdInh", icFirst)
        .GroupsInheritance = GetSetting(App.EXEName, "Preferences", "GrpInh", icFirst)
        .UseLivePreview = CBool(GetSetting(App.EXEName, "Preferences", "UseLivePreview", 1))
        .EnableUndoRedo = CBool(GetSetting(App.EXEName, "Preferences", "EnableUR", 1))
        .ImgSpace = Val(GetSetting(App.EXEName, "Preferences", "ImgSpace", 7))
        .language = GetSetting(App.EXEName, "Preferences", "Language", "eng")
        
        .UseInstallMenus = CBool(GetSetting(App.EXEName, "Preferences", "UseInstallMenus", 1))
        .UseMapView = CBool(GetSetting(App.EXEName, "Preferences", "UseMapView", 1))
        .LockProjects = CBool(GetSetting(App.EXEName, "Preferences", "LockProjects", 0))
        .UseEasyActions = CBool(GetSetting(App.EXEName, "Preferences", "UseEasyActions", 0))
        ResizeEasyControls
        SetupEasyControls
        
        .CodePage = GetSetting(App.EXEName, "Preferences", "Codepage", "")
        If LenB(.CodePage) = 0 Then
            .CodePage = "1252"
            If LenB(QueryValue(HKEY_CLASSES_ROOT, "MIME\Database\Codepage\50001", "BodyCharset")) <> 0 Then
                .CodePage = "50001"
            End If
        End If
        
        With .ToolbarStyle
            '.Font.FontName = GetSetting(App.EXEName, "Preferences", "ToolbarStyle.Font.FontName", Me.Font.Name)
            .Font.FontBold = CBool(GetSetting(App.EXEName, "Preferences", "ToolbarStyle.Font.FontBold", True))
            '.Font.FontItalic = CBool(GetSetting(App.EXEName, "Preferences", "ToolbarStyle.Font.FontItalic", Me.Font.Italic))
            '.Font.FontSize = Val(GetSetting(App.EXEName, "Preferences", "ToolbarStyle.Font.FontSize", Me.Font.Size))
            '.Font.FontUnderline = CBool(GetSetting(App.EXEName, "Preferences", "ToolbarStyle.Font.FontUnderline", Me.Font.Underline))
            .Color = Val(GetSetting(App.EXEName, "Preferences", "ToolbarStyle.Color", &H80000012))
        End With
        
        With .ToolbarItemStyle
            '.Font.FontName = GetSetting(App.EXEName, "Preferences", "ToolbarItemStyle.Font.FontName", Me.Font.Name)
            .Font.FontBold = CBool(GetSetting(App.EXEName, "Preferences", "ToolbarItemStyle.Font.FontBold", True))
            '.Font.FontItalic = CBool(GetSetting(App.EXEName, "Preferences", "ToolbarItemStyle.Font.FontItalic", Me.Font.Italic))
            '.Font.FontSize = Val(GetSetting(App.EXEName, "Preferences", "ToolbarItemStyle.Font.FontSize", Me.Font.Size))
            '.Font.FontUnderline = CBool(GetSetting(App.EXEName, "Preferences", "ToolbarItemStyle.Font.FontUnderline", Me.Font.Underline))
            .Color = Val(GetSetting(App.EXEName, "Preferences", "ToolbarItemStyle.Color", &H80000012))
        End With
        
        With .GroupStyle
            '.Font.FontName = GetSetting(App.EXEName, "Preferences", "GroupStyle.Font.FontName", Me.Font.Name)
            .Font.FontBold = CBool(GetSetting(App.EXEName, "Preferences", "GroupStyle.Font.FontBold", True))
            '.Font.FontItalic = CBool(GetSetting(App.EXEName, "Preferences", "GroupStyle.Font.FontItalic", Me.Font.Italic))
            '.Font.FontSize = Val(GetSetting(App.EXEName, "Preferences", "GroupStyle.Font.FontSize", Me.Font.Size))
            '.Font.FontUnderline = CBool(GetSetting(App.EXEName, "Preferences", "GroupStyle.Font.FontUnderline", Me.Font.Underline))
            .Color = Val(GetSetting(App.EXEName, "Preferences", "GroupStyle.Color", &H80000012))
        End With
        
        With .CommandStyle
            '.Font.FontName = GetSetting(App.EXEName, "Preferences", "CommandStyle.Font.FontName", Me.Font.Name)
            .Font.FontBold = CBool(GetSetting(App.EXEName, "Preferences", "CommandStyle.Font.FontBold", Me.Font.Bold))
            '.Font.FontItalic = CBool(GetSetting(App.EXEName, "Preferences", "CommandStyle.Font.FontItalic", Me.Font.Italic))
            '.Font.FontSize = Val(GetSetting(App.EXEName, "Preferences", "CommandStyle.Font.FontSize", Me.Font.Size))
            '.Font.FontUnderline = CBool(GetSetting(App.EXEName, "Preferences", "CommandStyle.Font.FontUnderline", Me.Font.Underline))
            .Color = Val(GetSetting(App.EXEName, "Preferences", "CommandStyle.Color", &H80000012))
        End With
        
        .DisabledItem = Val(GetSetting(App.EXEName, "Preferences", "DisabledItem.Color", &H80000011))
        .BrokenLink = Val(GetSetting(App.EXEName, "Preferences", "BrokenLink.Color", vbRed))
        .NoCompileItem = Val(GetSetting(App.EXEName, "Preferences", "NoCompileItem.Color", &H80000013))
        
        .VerifyLinksOptions.VerifyExternalLinks = CBool(GetSetting(App.EXEName, "Preferences", "VerifyExternalLinks", True))
        .VerifyLinksOptions.VerifyOptions = Val(GetSetting(App.EXEName, "Preferences", "VerifyOptions", 0))
        
        .ShowCleanPreview = CBool(GetSetting(App.EXEName, "Preferences", "ShowCleanPreview", 0))
        .AutoShowCompileReport = CBool(GetSetting(App.EXEName, "Preferences", "AutoShowCompileReport", 1))
        .EnableUnicodeInput = CBool(GetSetting(App.EXEName, "Preferences", "EnableUnicodeInput", 0))
    End With
    
    DoUNICODE = (GetSetting("DMB", "Preferences", "DoUNICODE", 1) = 1)
    
End Sub

Private Sub SavePrgPrefs()
    
    With Preferences
        SaveSetting App.EXEName, "Preferences", "AutoRecover", Abs(.AutoRecover)
        SaveSetting App.EXEName, "Preferences", "OpenLastProject", Abs(.OpenLastProject)
        SaveSetting App.EXEName, "Preferences", "SepHeight", .SepHeight
        SaveSetting App.EXEName, "Preferences", "ShowNag", Abs(.ShowNag)
        SaveSetting App.EXEName, "Preferences", "ShowWarningAIE", Abs(.ShowWarningAddInEditor)
        SaveSetting App.EXEName, "Preferences", "ShowPPOnNewProject", Abs(.ShowPPOnNewProject)
        SaveSetting App.EXEName, "Preferences", "CmdInh", .CommandsInheritance
        SaveSetting App.EXEName, "Preferences", "GrpInh", .GroupsInheritance
        SaveSetting App.EXEName, "Preferences", "UseLivePreview", Abs(.UseLivePreview)
        SaveSetting App.EXEName, "Preferences", "EnableUR", Abs(.EnableUndoRedo)
        SaveSetting App.EXEName, "Preferences", "ImgSpace", .ImgSpace
        SaveSetting App.EXEName, "Preferences", "Language", .language
        SaveSetting App.EXEName, "Preferences", "UseInstallMenus", Abs(.UseInstallMenus)
        SaveSetting App.EXEName, "Preferences", "LockProjects", Abs(.LockProjects)
        SaveSetting App.EXEName, "Preferences", "UseEasyActions", Abs(.UseEasyActions)
        SaveSetting App.EXEName, "Preferences", "UseMapView", Abs(.UseMapView)
        SaveSetting App.EXEName, "Preferences", "Codepage", .CodePage
        
        SaveSetting App.EXEName, "Preferences", "ToolbarStyle.Font.FontBold", Abs(.ToolbarStyle.Font.FontBold)
        SaveSetting App.EXEName, "Preferences", "ToolbarStyle.Color", .ToolbarStyle.Color
        
        SaveSetting App.EXEName, "Preferences", "ToolbarItemStyle.Font.FontBold", Abs(.ToolbarItemStyle.Font.FontBold)
        SaveSetting App.EXEName, "Preferences", "ToolbarItemStyle.Color", .ToolbarItemStyle.Color
        
        SaveSetting App.EXEName, "Preferences", "GroupStyle.Font.FontBold", Abs(.GroupStyle.Font.FontBold)
        SaveSetting App.EXEName, "Preferences", "GroupStyle.Color", .ToolbarStyle.Color
        
        SaveSetting App.EXEName, "Preferences", "CommandStyle.Font.FontBold", Abs(.CommandStyle.Font.FontBold)
        SaveSetting App.EXEName, "Preferences", "CommandStyle.Color", .CommandStyle.Color
        
        SaveSetting App.EXEName, "Preferences", "DisabledItem.Color", .DisabledItem
        SaveSetting App.EXEName, "Preferences", "BrokenLink.Color", .BrokenLink
        SaveSetting App.EXEName, "Preferences", "NoCompileItem.Color", .NoCompileItem
        
        SaveSetting App.EXEName, "Preferences", "VerifyExternalLinks", Abs(.VerifyLinksOptions.VerifyExternalLinks)
        SaveSetting App.EXEName, "Preferences", "VerifyOptions", .VerifyLinksOptions.VerifyOptions
        
        SaveSetting App.EXEName, "Preferences", "ShowCleanPreview", Abs(.ShowCleanPreview)
        SaveSetting App.EXEName, "Preferences", "AutoShowCompileReport", Abs(.AutoShowCompileReport)
        SaveSetting App.EXEName, "Preferences", "EnableUnicodeInput", Abs(.EnableUnicodeInput)
    End With
    
    'tbMenu.SaveToolbar "Software\VB and VBA Program Settings\" + App.EXEName, "Toolbars", "tb1C"
    'tbMenu2.SaveToolbar "Software\VB and VBA Program Settings\" + App.EXEName, "Toolbars", "tb2C"
    'tbCmd.SaveToolbar "Software\VB and VBA Program Settings\" + App.EXEName, "Toolbars", "tb3C"
    
    'SaveSetting App.EXEName, "Toolbars", "tb1V",  mnuToolbarsStandard.Checked
    'SaveSetting App.EXEName, "Toolbars", "tb2V", mnuToolbarsMenu.Checked
    'SaveSetting App.EXEName, "Toolbars", "tb3V", mnuToolbarsTools.Checked
    
End Sub

Private Sub SetupSubclassing(scState As Boolean)

    If msgSubClass Is Nothing Then Set msgSubClass = New xfxSC
    
    frmMainHWND = Me.hwnd
    msgSubClass.SubClassHwnd Me.hwnd, scState

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    IsSHIFT = False

End Sub

Private Sub Form_Load()

10        On Error GoTo Form_Load_Error

20        AppPath = AddTrailingSlash(App.Path, "\")

30        InitDone = False
40        Me.Enabled = False
          
50        ConfigIE
60        SetTBIcons
          
70        ResellerID = GetSetting(App.EXEName, "RegInfo", "ResellerID", "")
80        If Not FolderExists(AppPath + "rinfo\" + ResellerID) Then ResellerID = ""
90        If ResellerID <> "" Then ResellerInfo = Split(LoadFile(AppPath + "rinfo\" + ResellerID + "\data.dat"), vbCrLf)
          
100       mnuRegister.Visible = IsDEMO
110       mnuBrowsers.Visible = False
120       mnuStyleOptions.Visible = False
130       mnuPPShortcuts.Visible = False
140       mnuTBEShortcuts.Visible = False
    #If DEVVER = 1 Then
150           mnuToolsDynAPI.Visible = True
160           mnuToolsSep05.Visible = True
              mnuHelpDynAPIDoc.Visible = True
              mnuHelpSep05.Visible = True
    #Else
170           mnuToolsDynAPI.Visible = False
180           mnuToolsSep05.Visible = False
              mnuHelpDynAPIDoc.Visible = False
              mnuHelpSep05.Visible = False
    #End If
          
200       If Not IsDebug Then Set xMenu = New CMenu
210       Set TipsSys = New CTips

230       LoadPrgPrefs
240       LoadLocalizedStrings
          
250       With TipsSys
260           .DialogTitle = GetLocalizedStr(836)
270           .RegKey = App.EXEName
280       End With
          
290       ChkFileAssociation
'300       If IsWin9x Then
'310           tmrCheckPCHealth.Enabled = True
'320           tmrCheckPCHealth_Timer
'330       End If

340       GetSystemCharset
350       SetupTempFolders

          ' DISABLE ALL SOUND FEATURES
360       tbCmd.Buttons("tbSound").Visible = False

370       If IsDEMO Then
380           frmNag.Show vbModeless, Me
390           If Preferences.language <> "eng" Then
400               frmNag.lblInfo.FontName = "MS Sans Serif"
410               frmNag.lblInfo.FontBold = False
420           End If
430       Else
440           If Preferences.ShowNag Then
450               frmNag.Show vbModeless, Me
460               If Preferences.language <> "eng" Then
470                   frmNag.lblInfo.FontName = "MS Sans Serif"
480                   frmNag.lblInfo.FontBold = False
490               End If
500           End If
510       End If

      '    If NagScreenIsVisible And IsWinXP Then
      '        Do
      '            DoEvents
      '        Loop While frmNag.tmrReveal.Enabled
      '    Else
      '        DoEvents
      '    End If
          wbMainPreview.Navigate "about:blank"
520       DoEvents
          
530       If IsDebug Then
540           vlmCtrl.Enabled = False
550           If Not IsInIDE Then
560               DisplayTip GetLocalizedStr(538), GetLocalizedStr(241) + vbCrLf + vbCrLf + GetLocalizedStr(242) + vbCrLf + GetLocalizedStr(243), False
570           End If
580       Else
590           SetupSubclassing True
600       End If
          
610       ReDim MenuCmds(0)
620       ReDim MenuGrps(0)
         
630       If Val(GetSetting(App.EXEName, "WinPos", "X")) = 0 Then
640           picSplit.Left = Width / 2
              picSplit2.Top = Height / 2
650           CenterForm Me
660       Else
670           Left = GetSetting(App.EXEName, "WinPos", "X")
680           Top = GetSetting(App.EXEName, "WinPos", "Y")
690           Width = GetSetting(App.EXEName, "WinPos", "W")
700           Height = GetSetting(App.EXEName, "WinPos", "H")
710           picSplit.Left = GetSetting(App.EXEName, "WinPos", "SH", Width / 2)
              picSplit2.Top = GetSetting(App.EXEName, "WinPos", "SV", Height / 2)
              
720           If Left + Width / 2 > Screen.Width Or Top + Height / 2 > Screen.Height Then
730               Left = Screen.Width / 2 - Width / 2
740               Top = Screen.Height / 2 - Height / 2
750           End If
760       End If
          
770       SimonFile = TempPath + "eegg.exe"
780       Set FloodPanel.PictureControl = picFlood
          
790       SetupCharset Me
          
810       If IsWinNT And Not IsDEMO Then TransferRegInfo2AllUsers
          
820       If Preferences.ShowNag Then
830           tmrInit_Timer
840       Else
850           tmrInit.Enabled = True
860       End If

870       On Error GoTo 0
880       Exit Sub

Form_Load_Error:

890       MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Form.frmMain.Form_Load" & vbCrLf & vbCrLf & "Please verify that your license information is valid"
          
End Sub

Private Sub SetupTempFolders()

    TempPath = GetTEMPPath + "DMB\"
    PreviewPath = TempPath + "Preview\"
    If App.PrevInstance Then
        StatesPath = TempPath + "States\" & Timer * 100 & "\"
    Else
        StatesPath = TempPath + "States\Default\"
    End If
    MkDir2 PreviewPath
    MkDir2 StatesPath

End Sub

Private Sub TransferRegInfo2AllUsers()

    Const rk1 = "S-1-5-21-1757981266-602609370-725345543-%id\Software\"
    Const rk2 = "VB and VBA Program Settings\DMB\RegInfo"
    
    Dim Users() As NTUser
    Dim i As Integer
    Dim k As String
    Dim s() As String
    Dim ss As String
    Dim j As Integer
    
    On Error GoTo ExitSub
    
    Users = GetNTUsers
    
    For i = 0 To UBound(Users)
        k = Replace(rk1, "%id", Users(i).id)
        If QueryValue(HKEY_USERS, k + "Microsoft\Windows NT\CurrentVersion\Winlogon", "BuildNumber") > 0 Then
            k = k + rk2
            
            ss = ""
            s = Split(k, "\")
            For j = 0 To UBound(s)
                ss = ss + s(j) + "\"
                If j > 1 Then CreateNewKey Left(ss, Len(ss) - 1), HKEY_USERS
            Next j
            
            SetKeyValue HKEY_USERS, k, "CacheData", GetSetting(App.EXEName, "RegInfo", "CacheData", "")
            SetKeyValue HKEY_USERS, k, "CacheSig01", GetSetting(App.EXEName, "RegInfo", "CacheSig01", "")
            SetKeyValue HKEY_USERS, k, "CacheSig02", GetSetting(App.EXEName, "RegInfo", "CacheSig02", "")
            SetKeyValue HKEY_USERS, k, "Company", GetSetting(App.EXEName, "RegInfo", "Company", "")
            SetKeyValue HKEY_USERS, k, "InstallPath", GetSetting(App.EXEName, "RegInfo", "InstallPath", "")
            SetKeyValue HKEY_USERS, k, "OrderNum", GetSetting(App.EXEName, "RegInfo", "OrderNum", "")
            SetKeyValue HKEY_USERS, k, "PreRegVer", GetSetting(App.EXEName, "RegInfo", "PreRegVer", "")
            SetKeyValue HKEY_USERS, k, "SerialNumber", GetSetting(App.EXEName, "RegInfo", "SerialNumber", "")
            SetKeyValue HKEY_USERS, k, "ServerResponse", GetSetting(App.EXEName, "RegInfo", "ServerResponse", "")
            SetKeyValue HKEY_USERS, k, "SubSystem", GetSetting(App.EXEName, "RegInfo", "SubSystem", "")
            SetKeyValue HKEY_USERS, k, "User", GetSetting(App.EXEName, "RegInfo", "User", "")
            SetKeyValue HKEY_USERS, k, "Version", GetSetting(App.EXEName, "RegInfo", "Version", "")
        End If
    Next i
    
ExitSub:

End Sub

Private Sub SetMenusLCType()

    If Preferences.UseInstallMenus Then
        mnuToolsInstallMenusA.Visible = False
        mnuToolsInstallMenus.Visible = True
    Else
        mnuToolsInstallMenusA.Visible = True
        mnuToolsInstallMenus.Visible = False
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MimicHoverButtonHover False

End Sub

Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    If data.GetFormat(vbCFFiles) Then
        LoadMenu data.Files(1)
    End If

End Sub

Private Sub Form_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    If data.GetFormat(vbCFFiles) Then
        If Right$(data.Files(1), 3) = "dmb" Then
            Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectNone
        End If
    Else
        Effect = vbDropEffectNone
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error Resume Next

    If Project.HasChanged Then
        If (Not NewMenu) Then
            If Project.HasChanged Then Cancel = 1
        End If
    End If
    
    If Cancel = 0 Then
        RestoreIESettings
        SavePrgPrefs
        
        UnLockFile Project.FileName
    
        Unload frmSelectiveCopyPaste
        Unload frmFind
        Unload frmFontDialog
        Unload frmAddInEditor
        If PreviewIsOn Then Unload frmPreview
        
        If IsHVVisible Then showHelp "%%CLOSE%%"
    
        If Not IsDebug Then SetupSubclassing False
    
        SaveWinPos
        SaveRecent
        
        On Error Resume Next
        If FileExists(SimonFile) Then Kill SimonFile
        
        CleanPreviewDir True
        CleanStatesDir
        CleanPresetsDirs True
    End If
    
End Sub

Private Sub SaveWinPos()

    If Not InitDone Then Exit Sub

    If WindowState = vbNormal Then
        SaveSetting App.EXEName, "WinPos", "X", Left
        SaveSetting App.EXEName, "WinPos", "Y", Top
        SaveSetting App.EXEName, "WinPos", "W", Width
        SaveSetting App.EXEName, "WinPos", "H", Height
        SaveSetting App.EXEName, "WinPos", "SH", picSplit.Left
        SaveSetting App.EXEName, "WinPos", "SV", picSplit2.Top
    End If

End Sub

Private Sub Form_Resize()

    tmrResize.Enabled = False
    
    DoResize
    
    tmrResize.Enabled = True
        
End Sub

Private Sub DoResize()

    Dim TopPos As Long
    Dim cTop As Long
    
    Dim fdl As Long
    Dim fdt As Long
    Dim fdW As Long
    Dim fdH As Long
    
    Dim tppx As Integer
    Dim tppy As Integer
    
    tppx = Screen.TwipsPerPixelX
    tppy = Screen.TwipsPerPixelY
    
    'On Error Resume Next
    
    IsResizing = True
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picSplit.Left < 176 * tppx Then picSplit.Left = 176 * tppx
    If (Width - picSplit.Left) < 244 * tppx Then picSplit.Left = Width - 244 * tppx
    
    If picSplit2.Top < 110 * tppy Then picSplit2.Top = 110 * tppy
    If (Height - picSplit2.Top) < 177 * tppy Then picSplit2.Top = Height - 177 * tppy
    
    wbMainPreview.Visible = Preferences.UseLivePreview
    picSplit2.Visible = Preferences.UseLivePreview
    sbLP_Close.Visible = Preferences.UseLivePreview
    sbLP_GroupMode.Visible = Preferences.UseLivePreview
    
    cTop = GetClientTop(Me.hwnd)
    
    With sbDummy
'        If WindowState = vbMaximized Then
'            .Align = vbAlignNone
'            DoEvents
'            .Top = Height - .Height - 700
'            .Width = Width - 175
'            DoEvents
'        Else
'            .Align = vbAlignBottom
'        End If
        .Panels("sbFlood").Width = .Width - 27 * tppx
        picFlood.Move 2 * tppx, .Top + 4 * tppy, .Width - 27 * tppx, .Height - 5 * tppy
    End With
    
    TopPos = (Abs(tbMenu.Visible) + Abs(tbMenu2.Visible) + Abs(tbCmd.Visible)) * tbMenu.Height + tppy

    If wbMainPreview.Visible Then
        tvMenus.Move tppx, TopPos + tppy, picSplit.Left - 2 * tppx, picSplit2.Top - TopPos - 28 * tppy
    Else
        tvMenus.Move tppx, TopPos + tppy, picSplit.Left - 2 * tppx, Height - TopPos - sbDummy.Height - cTop - 29 * tppy
    End If
    tvMapView.Move tvMenus.Left, tvMenus.Top, tvMenus.Width, tvMenus.Height
    picSplit2.Left = tvMenus.Left + 3 * tppx
    picSplit2.Width = tvMenus.Width - 47 * tppx
    
    UpdateTSTabsLabels TopPos
    
    fdl = picSplit.Left + picSplit.Width + 2 * tppx
    fdt = TopPos + 35 * tppy
    fdW = Width - picSplit.Left - picSplit.Width - 13 * tppx
    fdH = Height - TopPos - sbDummy.Height - cTop - 42 * tppy
    
    picSplit.Top = TopPos
    picSplit.Height = fdH + 30 * tppy
    
    lblDataTitle.Move fdl + 8 * tppx, TopPos + 8 * tppy
    
    ' Caption Field
    lblCaption.Move fdl + 8 * tppx, fdt, fdW - 16 * tppx
    txtCaption.Move fdl + 8 * tppx, fdt + lblCaption.Height + tppy, fdW - 16 * tppx
    
    ' Enabled Control
    chkEnabled.Move fdl + 8 * tppx, txtCaption.Top + txtCaption.Height + 8 * tppy, GetTextSize(chkEnabled.caption + "XXXXXXXX", Me)(1) * tppx
    chkCompile.Move chkEnabled.Left + chkEnabled.Width + 8 * tppx, chkEnabled.Top, chkEnabled.Width
    
    ' Events' Tabs
    tsCmdType.Move fdl + 8 * tppx, chkEnabled.Top + chkEnabled.Height + 8 * tppy, fdW - 16 * tppy, 170 * tppy
    
    ' Status Text
    lblStatus.Move fdl + 8 * tppx, tsCmdType.Top + tsCmdType.Height + 8 * tppy
    txtStatus.Move fdl + 8 * tppx, lblStatus.Top + lblStatus.Height + tppy, fdW - 16 * tppx
    
    ' Alignment
    'lblAlignment.Move fdl + 120, txtStatus.Top + txtStatus.Height + 120
    'icmbAlignment.Move fdl + 120, lblAlignment.Top + lblAlignment.Height + 15, fdW - 240
    
    ' Cmds Layout
    lblLayout.Move fdl + 8 * tppx, txtStatus.Top + txtStatus.Height + 8 * tppy
    opAlignmentStyle(0).Move fdl + 8 * tppx, lblLayout.Top + lblLayout.Height + tppy
    lblASVertical.Move opAlignmentStyle(0).Left + opAlignmentStyle(0).Width + 4 * tppx, opAlignmentStyle(0).Top + (opAlignmentStyle(0).Height - lblASVertical.Height) \ 2
    opAlignmentStyle(1).Move fdl + 8 * tppx, opAlignmentStyle(0).Top + opAlignmentStyle(0).Height + 3 * tppy
    lblASHorizontal.Move opAlignmentStyle(1).Left + opAlignmentStyle(1).Width + 4 * tppx, opAlignmentStyle(1).Top + (opAlignmentStyle(1).Height - lblASHorizontal.Height) \ 2
    
    frameEvent.Move tsCmdType.Left + tppx, tsCmdType.Top + 25 * tppy, tsCmdType.Width - 3 * tppx, tsCmdType.Height - 27 * tppy
    
    ' Action Type
    lblActionType.Move 4 * tppx, 4 * tppy, tsCmdType.Width - 10 * tppx
    cmbActionType.Move 4 * tppx, lblActionType.Top + lblActionType.Height + tppx, tsCmdType.Width - 10 * tppx

    ResizeURLControls
    ResizeEasyControls
    
    cmdFindTargetGroup.Move cmbTargetMenu.Width + 6 * tppx, cmbTargetMenu.Top
    
    If Preferences.UseLivePreview Then
        wbMainPreview.Move tvMenus.Left + 2 * tppx, picSplit2.Top + picSplit2.Height + 11 * tppy, tvMenus.Width - 4 * tppx
        wbMainPreview.Height = Height - TopPos - picSplit2.Top - sbDummy.Height + 4 * tppy
        sbLP_Close.Move wbMainPreview.Left + wbMainPreview.Width - sbLP_GroupMode.Width, wbMainPreview.Top - 17 * tppy
        sbLP_GroupMode.Move wbMainPreview.Left + wbMainPreview.Width - sbLP_GroupMode.Width - sbLP_Close.Width, wbMainPreview.Top - 17 * tppy
    End If
    
    If WindowState = vbNormal And InitDone Then SaveWinPos
    
    IsResizing = False
    
    Refresh

End Sub

Private Sub ResizeURLControls()

    Dim tppx As Integer
    Dim tppy As Integer
    
    tppx = Screen.TwipsPerPixelX
    tppy = Screen.TwipsPerPixelY

    ' Action
    With lblActionName
        .Move 4 * tppx, cmbActionType.Top + cmbActionType.Height + 13 * tppy
        cmbTargetMenu.Move 4 * tppx, .Top + .Height + tppy, tsCmdType.Width - (13 * tppx + cmdBrowse.Width)
        
        If cmbActionType.ListIndex = 3 Then
            txtURL.Move 4 * tppx, .Top + .Height + tppy, tsCmdType.Width - cmdBrowse.Width - (100 * tppx)
        Else
            txtURL.Move 4 * tppx, .Top + .Height + tppy, tsCmdType.Width - cmdBrowse.Width - (67 * tppx)
        End If
        
        lblAlignment.Move 4 * tppx, .Top + 49 * tppy
        icmbAlignment.Move 4 * tppx, .Top + 63 * tppy, cmbActionType.Width
    End With
    With cmdBrowse
        .Move txtURL.Left + txtURL.Width + 4 * tppx, cmbTargetMenu.Top
        cmdBookmark.Move .Left + .Width + 4 * tppx, txtURL.Top
        cmdTargetFrame.Move cmdBookmark.Left + cmdBookmark.Width, txtURL.Top
        cmdWinParams.Move .Left + .Width + 4 * tppx, txtURL.Top, tsCmdType.Width - (.Left + .Width + 10 * tppx)
    End With

End Sub

Private Sub icmbAlignment_Click()

    If IsUpdating Then Exit Sub
    DontRefreshMap = True
    If IsGroup(GetID) Then
        If UBound(Project.Toolbars) = 0 Then
            DisplayTip GetLocalizedStr(688), GetLocalizedStr(689)
        End If
    End If
    UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(115)
    
End Sub

#If DEVVER = 1 Then
Private Sub MenuDelDynAPITemplate()

    Dim g As Integer
    Dim nNode As Node
    
    On Error Resume Next
    
    For g = 1 To UBound(MenuGrps)
        If MenuGrps(g).IsTemplate Then
            LastItemFP = ""
            SelectItem tvMenus.Nodes("G" & g)
            RemoveItem True, True
            Exit For
        End If
    Next g
    
    For Each nNode In tvMapView.Nodes
        If Project.Toolbars(ToolbarIndexByKey(nNode.key)).IsTemplate Then
            nNode.Selected = True
            IsTBMapSel = True
            RemoveItem True, True
            Exit For
        End If
    Next nNode
    
    RefreshMap

End Sub

Private Sub MenuAddDynAPITemplate()

    Dim tGrp As MenuGrp
    Dim tCmd As MenuCmd
    Dim g As Integer
    
    For g = 1 To UBound(MenuGrps)
        If MenuGrps(g).Name = "DynAPI_Template" Then
            MsgBox GetLocalizedStr(837), vbInformation + vbOKOnly, GetLocalizedStr(838)
            Exit Sub
        End If
    Next g

    tGrp = TemplateGroup
    tGrp.Name = "DynAPI_Template"
    tGrp.caption = "DynAPI_Template"
    tGrp.IsTemplate = True
    AddMenuGroup GetGrpParams(tGrp)
    
    SelectItem tvMenus.SelectedItem, True
    
    tCmd = TemplateCommand
    tCmd.Name = "DynAPICmd"
    tCmd.caption = "DynAPICmd"
    tCmd.Actions.onclick.url = "%%URL%%"
    AddMenuCommand GetCmdParams(tCmd)
    
    tCmd.Name = "[SEP]"
    tCmd.nTextColor = vbBlack
    tCmd.nBackColor = -2
    AddMenuCommand GetCmdParams(tCmd)
    
    AddToolbar
    With Project.Toolbars(UBound(Project.Toolbars))
        .IsTemplate = True
        .Name = "DynAPI_TBTemplate"
        ReDim .Groups(1)
        .Groups(1) = tGrp.Name
    End With
    
    RefreshMap

End Sub
#End If

Private Sub UpdateTSTabsLabels(Optional TopPos As Long)

    If TopPos <> 0 Then
        lblTSViewsNormal.Top = TopPos + tvMenus.Height + 90
        lblTSViewsMap.Top = lblTSViewsNormal.Top
    End If
    lblTSViewsNormal.Font.Bold = SelView = svcNormal
    lblTSViewsMap.Font.Bold = SelView = svcMap
    
    lblTSViewsNormal.Left = 135 + 30 * Abs(SelView = svcNormal)
    lblTSViewsMap.Left = 1110 + 30 * Abs(SelView = svcMap)

End Sub

Private Sub icmbEasyAlignment_Click()

    If IsUpdating Then Exit Sub
    
    Dim i As Integer
    
    For i = 1 To icmbEasyAlignment.ComboItems.Count
        If icmbEasyAlignment.ComboItems(i).Selected Then
            icmbAlignment.ComboItems(i).Selected = True
            icmbAlignment_Click
            Exit For
        End If
    Next i

End Sub

Private Sub lblTSViewsMap_Click()

    On Error Resume Next

    If SelView = svcMap Then Exit Sub

    InMapMode = True
    KeepExpansions = True
    RefreshMap True
    tvMapView.ZOrder 0
    SetCtrlFocus tvMapView

    SelView = svcMap
    Refresh
    
    UpdateTSTabsLabels
    
    UpdateControls

End Sub

Private Sub lblTSViewsNormal_Click()

    On Error Resume Next

    If SelView = svcNormal Then Exit Sub
    
    SaveExpansions

    tvMenus.ZOrder 0
    SetCtrlFocus tvMenus
    InMapMode = False
    tvMapView.Nodes.Clear
    
    IsTBMapSel = False
    
    SelView = svcNormal
    Refresh
    
    UpdateTSTabsLabels
    
    UpdateControls

End Sub

Private Sub ShowPreferences()

    Dim LastURState As Boolean
    Dim OldLang As String
    Dim OldViewMode As Boolean

    OldLang = Preferences.language
    LastURState = Preferences.EnableUndoRedo
    OldViewMode = Preferences.UseMapView
    
    frmPreferences.Show vbModal
    
    If Preferences.EnableUndoRedo And LastURState = False Then
        ResetUndoStates
    End If
    
    If OldViewMode = False And Preferences.UseMapView Then
        lblTSViewsMap_Click
    Else
        If OldViewMode = True And Not Preferences.UseMapView Then
            lblTSViewsNormal_Click
        End If
    End If
    
    If OldLang <> Preferences.language Then
        LoadLocalizedStrings
        LocalizeUI
        ChkFileAssociation True
        SetCombosData
        RefreshMap
    End If
    
    SetupEasyControls
    
    SetMenusLCType
    UpdateControls
    
    If InMapMode Then RefreshMap

End Sub

Private Sub SetupEasyControls()

    With picEasy
        .ZOrder 0
        .BorderStyle = 0
        .BorderStyle = 0
        
        .Visible = Preferences.UseEasyActions
    End With
    
    With chkEasyEnableNewWindow
        .Top = sbEasyNewWindow.Top + sbEasyNewWindow.Height / 2 - .Height / 2
    End With

End Sub

Private Function ShortenName(ByVal n As String) As String

    Dim s() As String
    Dim i As Integer
    Dim k As Integer
    
    n = Replace(n, "[", "")
    n = Replace(n, "]", "")
    n = Replace(n, "_", "")
    
    s = Split(n, " ")
    n = ""
    If UBound(s) > 1 Then
        k = 2
    Else
        k = 3
    End If
    
    For i = 0 To UBound(s)
        n = n + Left(s(i), k)
    Next i
    
    ShortenName = n

End Function

Private Sub MenuAddSubGroup()

    Dim SelId As Integer
    
    FinishRenaming
    IsRenaming = False

    If Not InMapMode Then Exit Sub
    SelId = GetID
    
    KeepExpansions = False

    TemplateGroup.Actions.onmouseover.Type = atcCascade
    AddMenuGroup , InMapMode
    TemplateGroup.Actions.onmouseover.Type = atcNone
    
    If InMapMode Then
        MenuGrps(UBound(MenuGrps)).Name = FixItemName(GetSecuenceName(True, "SubGrp_" & ShortenName(tvMapView.SelectedItem.Text)))
        With MenuCmds(SelId).Actions.onmouseover
            .Type = atcCascade
            .TargetMenu = GetID
        End With
        tvMenus.SelectedItem.Text = MenuGrps(UBound(MenuGrps)).Name
        
        LastSelNode = ""
        RefreshMap
        
        SelId = MenuCmds(SelId).parent
        With MenuGrps(GetID)
            .Actions.onmouseover.TargetMenu = GetID
            .AlignmentStyle = MenuGrps(SelId).AlignmentStyle
            .BackImage = MenuGrps(SelId).BackImage
            .bColor = MenuGrps(SelId).bColor
            .BorderStyle = MenuGrps(SelId).BorderStyle
            .CmdsFXhColor = MenuGrps(SelId).CmdsFXhColor
            .CmdsFXnColor = MenuGrps(SelId).CmdsFXnColor
            .CmdsFXNormal = MenuGrps(SelId).CmdsFXNormal
            .CmdsFXOver = MenuGrps(SelId).CmdsFXOver
            .CmdsFXSize = MenuGrps(SelId).CmdsFXSize
            .CmdsMarginX = MenuGrps(SelId).CmdsMarginX
            .CmdsMarginY = MenuGrps(SelId).CmdsMarginY
            .ContentsMarginH = MenuGrps(SelId).ContentsMarginH
            .ContentsMarginV = MenuGrps(SelId).ContentsMarginV
            .Corners = MenuGrps(SelId).Corners
            .CornersImages = MenuGrps(SelId).CornersImages
            .DefHoverFont = MenuGrps(SelId).DefHoverFont
            .DefNormalFont = MenuGrps(SelId).DefNormalFont
            .DropShadowColor = MenuGrps(SelId).DropShadowColor
            .DropShadowSize = MenuGrps(SelId).DropShadowSize
            .frameBorder = MenuGrps(SelId).frameBorder
            .hBackColor = MenuGrps(SelId).hBackColor
            .hTextColor = MenuGrps(SelId).hTextColor
            .iCursor = MenuGrps(SelId).iCursor
            .Image = MenuGrps(SelId).Image
            .Leading = MenuGrps(SelId).Leading
            .tbiLeftImage = MenuGrps(SelId).tbiLeftImage
            .nBackColor = MenuGrps(SelId).nBackColor
            .nTextColor = MenuGrps(SelId).nTextColor
            .tbiRightImage = MenuGrps(SelId).tbiRightImage
            .Transparency = MenuGrps(SelId).Transparency
            .tbiBackImage = MenuGrps(SelId).tbiBackImage
        End With
    End If
    
    SaveState GetLocalizedStr(760) + " " + tvMapView.SelectedItem.Text
    
    'If IsDebug Then SetupMenuMenu tvMenus.SelectedItem
    
    MenuAddCommand

End Sub

Private Sub AddToolbar()

    Dim nName As String
    Dim nNode As Node
    
    FinishRenaming
    IsRenaming = False

    ReDim Preserve Project.Toolbars(UBound(Project.Toolbars) + 1)
    With Project.Toolbars(UBound(Project.Toolbars))
        nName = GetTBSecuenceName("Untitled")
        .Name = nName
        .IsTemplate = False
        ReDim .Groups(0)
        .Compile = True
    End With
    
    RefreshMap
    
    For Each nNode In tvMapView.Nodes
        If nNode.Text = nName Then
            nNode.Selected = True
            nNode.EnsureVisible
            LastSelNode = nNode.tag
            Exit For
        End If
    Next nNode
    
    KeepExpansions = False
    UpdateControls
    
    IsTBMapSel = True
    tvMapView.StartLabelEdit
    
    SaveState GetLocalizedStr(809) + " " + tvMapView.SelectedItem.Text
    
    If IsDebug Then SetupMenuMenu tvMapView.SelectedItem

End Sub

Private Sub ShowLengthDlg()

    frmSepPer.Show vbModal
    
    UpdateLivePreview

End Sub

Private Sub ToolsPublish()

    If SomethingIsWrong(False, False) Or _
        InvalidItemsNames Or _
        DuplicatedItemsNames Or _
        InvalidItemsData Or _
        MissingParameters Then
        Exit Sub
    End If

    frmFTPPublishing.Show vbModal

End Sub

Friend Sub ToolsCompile()

    Dim OriginalProject As ProjectDef
    Dim dc As Integer
    Dim i As Integer
    Dim j As Integer
    
    Dim sTB As Integer
    Dim sCmd As Integer
    Dim sGrp As Integer
    Dim sValid As Boolean
    
    Dim k As Integer
    Dim jsfn() As String
    
    Dim IsMultiProjects As Boolean
    
    IsMultiProjects = UBound(Project.SecondaryProjects) > 0
    If IsMultiProjects Then
        If Project.CodeOptimization = cocAggressive Then
            MsgBox "The 'max' Code Optimization setting is not supported when working with Secondary Projects", vbInformation + vbOKOnly, GetLocalizedStr(569)
            Exit Sub
        End If
    End If
    
    #If DEVVER = 1 Then
        If mnuToolsDynAPI.Checked Then
            If Project.CodeOptimization = cocAggressive Then
                MsgBox "The 'max' Code Optimization setting is not supported when using the DynAPI", vbInformation + vbOKOnly, GetLocalizedStr(569)
                Exit Sub
            End If
        End If
    #End If
    
    DoEvents
    
    OriginalProject = Project
    If SomethingIsWrong(False, False) Or _
        InvalidItemsNames Or _
        DuplicatedItemsNames Or _
        InvalidItemsData Or _
        MissingParameters Then
        GoTo AbortCompile
    End If
    
    If Project.UserConfigs(Project.DefaultConfig).Type = ctcLocal Then
        DisplayTip GetLocalizedStr(684), GetLocalizedStr(685)
    End If
    
    If DoCompile Then
        ' START Secondary projects support ********************************
        If IsMultiProjects Then
            ReDim jsfn(1)
            jsfn(1) = OriginalProject.JSFileName
            
            SaveMenu False
            
            For i = 1 To UBound(Project.SecondaryProjects)
                sTB = sTB + UBound(Project.Toolbars)
                For j = 1 To UBound(Project.Toolbars)
                    sGrp = sGrp + UBound(Project.Toolbars(j).Groups)
                Next j
                sCmd = sCmd + UBound(MenuCmds)
                LoadMenu OriginalProject.SecondaryProjects(i)
                If Not FileExists(OriginalProject.SecondaryProjects(i)) Then Exit For

                With Project
                    .CodeOptimization = OriginalProject.CodeOptimization
                    
                    k = 0
ReStart:
                    For j = 1 To UBound(jsfn)
                        If jsfn(j) = .JSFileName Then
                            .JSFileName = .JSFileName & "P" & (i + k)
                            k = k + 1
                            SaveMenu False
                            GoTo ReStart
                        End If
                    Next j
                    ReDim Preserve jsfn(UBound(jsfn) + 1)
                    jsfn(UBound(jsfn)) = .JSFileName

                    sValid = False
                    
                    ' Test code *******************************************
                    .UserConfigs = OriginalProject.UserConfigs
                    .DefaultConfig = OriginalProject.DefaultConfig
                    ' *****************************************************
                    
                    If UBound(.UserConfigs) = UBound(OriginalProject.UserConfigs) Then
                        j = .DefaultConfig
                        dc = OriginalProject.DefaultConfig
                        sValid = True
                        sValid = sValid And (.UserConfigs(j).CompiledPath = OriginalProject.UserConfigs(dc).CompiledPath)
                        sValid = sValid And (.UserConfigs(j).Frames.UseFrames = OriginalProject.UserConfigs(dc).Frames.UseFrames)
                        sValid = sValid And (.UserConfigs(j).ImagesPath = OriginalProject.UserConfigs(dc).ImagesPath)
                        sValid = sValid And (.UserConfigs(j).RootWeb = OriginalProject.UserConfigs(dc).RootWeb)
                        sValid = sValid And (.UserConfigs(j).Type = OriginalProject.UserConfigs(dc).Type)
                        If OriginalProject.UserConfigs(dc).Type = ctcRemote Then
                            dc = GetConfigID(OriginalProject.UserConfigs(dc).LocalInfo4RemoteConfig)
                            j = GetConfigID(.UserConfigs(j).LocalInfo4RemoteConfig)
                            sValid = sValid And (.UserConfigs(j).CompiledPath = OriginalProject.UserConfigs(dc).CompiledPath)
                            sValid = sValid And (.UserConfigs(j).Frames.UseFrames = OriginalProject.UserConfigs(dc).Frames.UseFrames)
                            sValid = sValid And (.UserConfigs(j).ImagesPath = OriginalProject.UserConfigs(dc).ImagesPath)
                            sValid = sValid And (.UserConfigs(j).RootWeb = OriginalProject.UserConfigs(dc).RootWeb)
                            sValid = sValid And (.UserConfigs(j).Type = OriginalProject.UserConfigs(dc).Type)
                        End If
                        sValid = sValid And Not .UserConfigs(j).Frames.UseFrames
                    End If
                    If sValid Then
                        DoCompile , , , sTB, sGrp, sCmd
                    Else
                        MsgBox GetLocalizedStr(840) + " '" + .Name + "' " + GetLocalizedStr(841), vbInformation + vbOKOnly, GetLocalizedStr(569)
                        Exit For
                    End If
                End With
            Next i
            LoadMenu OriginalProject.FileName
        End If
        ' END   Secondary projects support ********************************
        If Preferences.AutoShowCompileReport Then frmCompilationReport.Show vbModal
    End If
    
AbortCompile:
    Project.CodeOptimization = OriginalProject.CodeOptimization
    Project.DefaultConfig = OriginalProject.DefaultConfig

End Sub

Private Function MissingParameters() As Boolean

    Dim i As Integer
    Dim MissingParam As Boolean
    
    If LenB(Project.AddIn.Name) <> 0 Then
        For i = 1 To UBound(params)
            With params(i)
                If .Required And LenB(.Value) = 0 Then
                    MissingParam = True
                    MsgBox GetLocalizedStr(246) + " " + Project.AddIn.Name + _
                            " " + GetLocalizedStr(247) + vbCrLf + GetLocalizedStr(248), vbInformation + vbOKOnly, GetLocalizedStr(541)
                End If
            End With
        Next i
    End If
    
    MissingParameters = MissingParam

End Function

Private Sub mnuBrowsersList_Click(Index As Integer)

    SaveSetting App.EXEName, "Browsers", "Default", Index + 1
    SetDefBrowserIcon
    DoEvents
    ShowPreview

End Sub

Private Sub mnuBrowsersSetDefBrowser_Click()

    ShowBrowsersDlg

End Sub

Private Sub mnuEditCopy_Click()

    SelectiveLauncher = slcSelCopy
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuEditDelete_Click()

    RemoveItem

End Sub

Private Sub mnuEditFind_Click()

    DoFind

End Sub

Private Sub mnuEditFindNext_Click()

    frmFind.DoFind

End Sub

Private Sub mnuEditFindReplace_Click()

    With frmFind
        .Show vbModeless, Me
        .SwitchToReplaceMode
    End With

End Sub

Private Sub mnuEditPaste_Click()

    SelectiveLauncher = slcSelPaste
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuEditPreferences_Click()
    
    ShowPreferences

End Sub

Private Sub mnuEditRedo_Click()

    DoRedo

End Sub

Private Sub mnuEditRename_Click()

    If InMapMode Then
        tvMapView.StartLabelEdit
    Else
        tvMenus.StartLabelEdit
    End If

End Sub

Private Sub mnuEditUndo_Click()

    DoUndo

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub mnuFileExportHTML_Click()
   
    DisplayTip "Saving the project as HTML", "The Save As HTML option has been specifically designed to let you generate plain HTML versions of your menus." + vbCrLf + _
                "This type of output is often useful to generate sitemaps." + vbCrLf + vbCrLf + _
                "Note that in order to actually generate the menus as pulldown menus you should use the Tools->Compile option."
    frmExportHTML.Show vbModal

End Sub

Private Sub mnuFileExportSitemap_Click()

    Dim FileName As String
    
    On Error Resume Next
    
    With cDlg
        .DialogTitle = "Select the folder where you want to save the sitemap"
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNNoReadOnlyReturn
        .filter = "XML Document|*.xml"
        .InitDir = GetRealLocal.RootWeb
        .FileName = GetRealLocal.RootWeb + "sitemap.xml"
        .CancelError = True
        .ShowSave
        If Err.number > 0 Or LenB(.FileName) = 0 Then Exit Sub
        FileName = .FileName
        If LCase(GetFileExtension(FileName)) <> "xml" Then FileName = FileName + ".xml"
    End With

    modExport.ExportAsSitemap FileName

End Sub

Private Sub mnuFileNewEmpty_Click()

    ProjectPropertiesPage = pppcGeneral
    NewEmptyProject

End Sub

Private Sub mnuFileNewFromDir_Click()

    If NewMenu Then
        IsLoadingProject = True
        frmDirSS.Show vbModal
        IsLoadingProject = False
        
        LastItemFP = ""
        If UBound(MenuGrps) > 0 Then SelectItem tvMenus.Nodes(1)
        
        RefreshMap
    End If

End Sub

Private Sub mnuFileNewFromPreset_Click()

    If NewMenu Then
        PresetWorkingMode = pwmNormal
        frmPresetsManager.Show vbModal
        If Preferences.ShowPPOnNewProject Then ShowProjectProperties
    End If

End Sub

Private Sub mnuFileNewFromROR_Click()

    RunShellExecute "open", "ror2dmb.exe", "", Long2Short(AppPath), 1

End Sub

Private Sub mnuFileNewFromTXT_Click()

    frmImportFromTXT.Show vbModal
    LastItemFP = ""
    If UBound(MenuGrps) > 0 Then SelectItem tvMenus.Nodes(1)
    
    RefreshMap

End Sub

Private Sub mnuFileNewFromWizard_Click()
    
    RunShellExecute "open", "dmbwizard.exe", "", Long2Short(AppPath), 1
    
End Sub

Private Sub mnuFileOpen_Click()

    LoadMenu

End Sub

Private Sub mnuFileOpenRecentOP_Click()

    With mnuFileOpenRecentOP
        .Checked = Not .Checked
        Preferences.OpenLastProject = .Checked
    End With

End Sub

Private Sub mnuFileOpenRecentR_Click(Index As Integer)

    OpenRecent Index + 1

End Sub

Private Sub mnuFileProjProp_Click()
    
    ProjectPropertiesPage = pppcGeneral
    ShowProjectProperties

End Sub

Private Sub mnuFileSave_Click()

    SaveMenu

End Sub

Private Sub mnuFileSaveAs_Click()

    FileSaveAs

End Sub

Private Sub mnuFileSaveAsPreset_Click()

    If Not CreateToolbar Then
        MsgBox GetLocalizedStr(842), vbInformation + vbOKOnly, GetLocalizedStr(843)
    Else
        frmPresetCreate.Show vbModal
    End If

End Sub

Private Sub mnuFileSubmitPreset_Click()

    If IsDEMO Then
        MsgBox GetLocalizedStr(899), vbInformation + vbOKOnly, GetLocalizedStr(900)
    Else
        PresetWorkingMode = pwmSubmit
        frmPresetsManager.Show vbModal
    End If

End Sub

Private Sub mnuHelpAbout_Click()

    frmAbout.Show vbModal

End Sub

Private Sub mnuHelpContents_Click()

    showHelp "introduction.htm"

End Sub

Private Sub mnuHelpDLPDF_Click()

    RunShellExecute "Open", "https://xfx.net/utilities/dmbuilder/download.php#PDFDoc", 0, 0, 0

End Sub

Private Sub mnuHelpDynAPIDoc_Click()

    RunShellExecute "Open", "https://xfx.net/utilities/dmbuilderde/help/index.html", 0, 0, 0

End Sub

Private Sub mnuHelpPosMenus_Click()

    showHelp "thetutorial_ptb_01.htm"

End Sub

Private Sub mnuHelpSearch_Click()

    showHelp "introduction.htm"
    showHelp "%%SEARCH%%"

End Sub

Private Sub mnuHelpTutorialsHS_Click()

    showHelp "thetutorial_ntb_01.htm"

End Sub

Private Sub mnuHelpTutorialsVideo_Click()

    RunShellExecute "Open", "https://xfx.net/utilities/dmbuilder/content/videos/tb/index.html", 0, 0, 0

End Sub

Private Sub mnuHelpTutotialsTB_Click()

    showHelp "thetutorial_tb_01.htm"

End Sub

Private Sub mnuHelpUpgrade_Click()

    #If DEVVER = 1 Then
        RunShellExecute "Open", "https://xfx.net/utilities/dmbuilder/download.htm", 0, 0, 0
    #Else
        DoUpgrade
    #End If

End Sub

Private Sub mnuHelpXFXFAQ_Click()

    RunShellExecute "Open", "https://xfx.net/utilities/dmbuilder/faq/index.html", 0, 0, 0

End Sub

Private Sub mnuHelpXFXHomePage_Click()

    RunShellExecute "Open", "https://xfx.net/index.html", 0, 0, 0

End Sub

Private Sub mnuHelpXFXNews_Click()

    RunShellExecute "Open", "https://xfx.net/utilities/dmbuilder/news.htm", 0, 0, 0

End Sub

Private Sub mnuHelpXFXPublicForum_Click()

    RunShellExecute "Open", "https://xfx.net/uboards/uboard_dmb.htm", 0, 0, 0

End Sub

Private Sub mnuHelpXFXSupport_Click()

    Dim ORDERNUMBER As String
'    Dim i As Integer
'    Dim sStr As String
'    Dim eMail As String
    
    If IsDEMO Then
        ORDERNUMBER = "DEMO"
    Else
        ORDERNUMBER = GetSetting(App.EXEName, "RegInfo", "OrderNum", "DEMO")
    End If
    
'    i = 0
'    Do
'        sStr = EnumSubKeys(HKEY_CURRENT_USER, "Software\Microsoft\Internet Account Manager\Accounts", i)
'        If sStr = "" Then Exit Do
'        eMail = QueryValue(HKEY_CURRENT_USER, "Software\Microsoft\Internet Account Manager\Accounts" + sStr, "")
'    Loop

    RunShellExecute "Open", "https://xfx.net/ss.php?un=" + USER + "&on=" + ORDERNUMBER + "&ver=" + DMBVersion + "&em=" + "", 0, 0, 0

End Sub

Private Sub mnuMenuAddCommand_Click()

    MenuAddCommand

End Sub

Private Sub mnuMenuAddGroup_Click()

    MenuAddGroup

End Sub

Private Sub mnuMenuAddSeparator_Click()

    MenuAddSeparator

End Sub

Private Sub mnuMenuAddSubGroup_Click()

    MenuAddSubGroup

End Sub

Private Sub mnuMenuAddToolbar_Click()

    AddToolbar

End Sub

Private Sub mnuMenuColor_Click()

    SelectiveLauncher = slcSelColor
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuMenuCopy_Click()

    mnuEditCopy_Click

End Sub

Private Sub mnuMenuCursor_Click()
    
    SelectiveLauncher = slcSelCursor
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuMenuDelete_Click()

    mnuEditDelete_Click

End Sub

Private Sub mnuMenuFont_Click()

    SelectiveLauncher = slcSelFont
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuMenuImage_Click()
    
    SelectiveLauncher = slcSelImage
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuMenuLength_Click()

    SelectiveLauncher = slcSelSepLen
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuMenuMargins_Click()

    SelectiveLauncher = slcSelMargin
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuMenuPaste_Click()

    mnuEditPaste_Click

End Sub

Private Sub mnuMenuRemoveToolbar_Click()

    mnuEditDelete_Click

End Sub

Private Sub mnuMenuRename_Click()

    mnuEditRename_Click

End Sub

Private Sub mnuMenuSelFX_Click()

    SelectiveLauncher = slcSelSelFX
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuMenuSFX_Click()

    SelectiveLauncher = slcSelSFX
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuMenuToolbarProperties_Click()
    
    SelectiveLauncher = slcSelTBEditor
    tmrLauncher.Enabled = True
    
End Sub

Private Sub mnuPPShortcutsAdvanced_Click()

    ProjectPropertiesPage = pppcAdvanced
    SelectiveLauncher = slcProjectProperties
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuPPShortcutsConfig_Click()

    ProjectPropertiesPage = pppcConfig
    SelectiveLauncher = slcProjectProperties
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuPPShortcutsGeneral_Click()

    ProjectPropertiesPage = pppcGeneral
    SelectiveLauncher = slcProjectProperties
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuPPShortcutsGlobal_Click()

    ProjectPropertiesPage = pppcGlobal
    SelectiveLauncher = slcProjectProperties
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuRegisterBuy_Click()

    If IsDEMO Then
        frmPurchase.Show vbModal
        'RunShellExecute "Open", "https://xfx.net/utilities/dmbuilder/register.htm", 0, 0, 0
    End If

End Sub

Private Sub mnuRegisterUnlock_Click()

    If IsDEMO Then DoUnlock

End Sub

Private Sub mnuStyleOptionsOP_Click(Index As Integer)

    Dim i As Integer
    Dim sbOP As Control
    Dim c As String
    Dim t As Integer
    
    On Error Resume Next
    
    mnuStyleOptionsOPAdvanced.Checked = (Index = -1)
    For i = 0 To mnuStyleOptionsOP.Count - 1
        mnuStyleOptionsOP(i).Checked = (i = Index)
    Next i
    
    Set sbOP = Screen.ActiveForm.sbApplyOptions
    With sbOP
        If Index = -1 Then
            c = ""
            t = UBound(dmbClipboard.CustomSel)
            For i = 1 To t
                If dmbClipboard.ObjSrc = docCommand Then
                    c = c + NiceCmdCaption(GetIDByName(dmbClipboard.CustomSel(i)))
                Else
                    c = c + NiceGrpCaption(GetIDByName(dmbClipboard.CustomSel(i)))
                End If
                Select Case i
                    Case t - 1
                        c = c + " and "
                    Case Is < t - 1
                        c = c + ", "
                End Select
            Next i
            c = Replace(GetLocalizedStr(920), "%%ITEMNAME%%", c)
        Else
            c = mnuStyleOptionsOP(Index).caption
        End If
        .ToolTipText = c
        .caption = "     " + c
        Do While SetCtrlWidth(sbOP) > .Width + (18 * 30)
            .caption = Left(.caption, Len(.caption) - 2) + "…"
        Loop
    End With

End Sub

Private Sub mnuStyleOptionsOPAdvanced_Click()

    Dim i As Integer
    
    On Error Resume Next

    With dmbClipboard
        If IsGroup(tvMenus.SelectedItem.key) Then
            .GrpContents = MenuGrps(GetID)
            .ObjSrc = docGroup
        Else
            If IsCommand(tvMenus.SelectedItem.key) Then
                .CmdContents = MenuCmds(GetID)
                .ObjSrc = docCommand
            Else
                MsgBox "This feature is not supported on separators", vbInformation + vbOKOnly, "Advanced Scope Options"
            End If
        End If
    End With
    frmSelItems.Show vbModal
    
    i = UBound(dmbClipboard.CustomSel)
    If Err.number > 0 Then Exit Sub
    
    mnuStyleOptionsOP_Click -1

End Sub

Private Sub mnuTBEShortcutsAdvanced_Click()

    TBEPage = tbepcAdvanced
    SelectiveLauncher = slcSelTBEditor
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuTBEShortcutsAppearance_Click()

    TBEPage = tbepcAppearance
    SelectiveLauncher = slcSelTBEditor
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuTBEShortcutsEffects_Click()

    TBEPage = tbepcEffects
    SelectiveLauncher = slcSelTBEditor
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuTBEShortcutsGeneral_Click()

    TBEPage = tbepcGeneral
    SelectiveLauncher = slcSelTBEditor
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuTBEShortcutsPositioning_Click()

    TBEPage = tbepcPositioning
    SelectiveLauncher = slcSelTBEditor
    tmrLauncher.Enabled = True

End Sub

Private Sub mnuToolsAddInEditor_Click()

    #If LITE = 1 Then
        ShowLITELImitationInfo 2
    #Else
        ShowAddInEditor
    #End If

End Sub

Private Sub mnuToolsApplyStyle_Click()

    PresetWorkingMode = pwmApplyStyle
    frmPresetsManager.Show vbModal

End Sub

Private Sub mnuToolsBrokenLinks_Click()

    LinkVerifyMode = spmcManual
    BrokenLinksReport

End Sub

Private Sub BrokenLinksReport()

    #If LITE = 1 Then
        ShowLITELImitationInfo 2
    #Else
        frmBrokenLinksReport.Show vbModal
        UpdateControls
        RefreshMap
    #End If

End Sub

Private Sub mnuToolsCompile_Click()

    ToolsCompile

End Sub

Private Sub mnuToolsDefaultConfig_Click()

    ReSetDefaultConfig

End Sub

Private Sub mnuToolsDynAPI_Click()

    #If DEVVER = 1 Then
    
    If IsLoadingProject Then Exit Sub
    mnuToolsDynAPI.Checked = Not mnuToolsDynAPI.Checked
    Project.GenDynAPI = mnuToolsDynAPI.Checked
    
    If Project.GenDynAPI Then
        MenuAddDynAPITemplate
    Else
        MenuDelDynAPITemplate
    End If
    
    #End If

End Sub

'Private Sub FileImport()
'
'    On Error GoTo ExitSub
'
'    If Not NewMenu(False) Then Exit Sub
'
'    With cDlg
'        .DialogTitle = GetLocalizedStr(250)
'        .filter = SupportedImportDocs
'        .CancelError = True
'        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
'        .ShowOpen
'        IsUpdating = True
'
'        IsLoadingProject = True
'        tvMenus.Visible = False
'        Me.Enabled = False
'
'        ImportProject .FileName
'        IsUpdating = False
'
'        If tvMenus.Nodes.count > 0 Then
'            With tvMenus.Nodes(1)
'                .Selected = True
'                .EnsureVisible
'            End With
'            UpdateLivePreview
'            DoEvents
'            SelectItem tvMenus.Nodes(1)
'            UpdateControls
'        End If
'    End With
'
'ExitSub:
'
'    IsLoadingProject = False
'    tvMenus.Visible = True
'    Me.Enabled = True
'
'    Exit Sub
'
'End Sub

Private Sub ShowProjectProperties()

    frmProjProp.Show vbModal
    
    If Project.HasChanged Then
        UpdateTitleBar
        UpdateTargetFrameCombo
        UpdateControls
        If InMapMode Then RefreshMap
    End If
    
    With GetRealLocal
        If LenB(.RootWeb) = 0 Or LenB(.CompiledPath) = 0 Or LenB(.ImagesPath) = 0 Then
            DisplayTip GetLocalizedStr(691), GetLocalizedStr(692)
        End If
    End With
    
End Sub

Friend Function NewMenu(Optional IsState As Boolean) As Boolean

    Dim Ans As Integer
    
    On Error GoTo NewMenu_Error

    If Not IsState Then
        If Project.HasChanged Then
            Ans = MsgBox(GetLocalizedStr(251) + " " + Project.Name + "?", vbQuestion + vbYesNoCancel, GetLocalizedStr(542))
            Select Case Ans
                Case vbYes
                    SaveMenu
                Case vbCancel
                    NewMenu = False
                    Exit Function
            End Select
        End If
    End If
    
    If FileIsLocked(Project.FileName) Then UnLockFile Project.FileName
    
    With Project
        .Name = "untitled"
        .AbsPath = ""
        .FileName = ""
        .AddIn.Name = ""
        .CodeOptimization = cocNormal
        .RemoveImageAutoPosCode = False
        .FX = 0
        .HasChanged = False
        
        ReDim .UserConfigs(0)
        With .UserConfigs(0)
            .Name = GetLocalizedStr(400)
            .Description = GetLocalizedStr(282)
            .Type = ctcCDROM
        
            .RootWeb = ""
            .CompiledPath = ""
            .ImagesPath = ""
            .OptmizePaths = False
        
            .HotSpotEditor.HotSpotsFile = ""
    
            .Frames.UseFrames = False
            .Frames.FramesFile = ""
        End With
        .DefaultConfig = 0
        
        ReDim .Toolbars(1)
        With .ToolBar
            .Name = GetLocalizedStr(400)
            .Alignment = tbacTopCenter
            .BackColor = &H808080
            
            .bOrder = 1
            .BorderColor = &H808080
            .BorderStyle = cfxcNone
            
            .CustX = 0
            .CustY = 0
            
            .FollowHScroll = False
            .FollowVScroll = False
            ReDim .Groups(0)
            .OffsetH = 0
            .OffsetV = 0
            .Image = ""
            .JustifyHotSpots = False
            .Spanning = tscAuto
            .Style = tscHorizonal
            .Separation = 1
            
            .DropShadowColor = &H999999
            .DropShadowSize = 0
            .Transparency = 0
            
            .Compile = True
        End With
        .Toolbars(1) = .ToolBar
        
'        .FTP.FTPAddress = ""
'        .FTP.Password = ""
'        .FTP.ProxyAddress = ""
'        .FTP.ProxyPort = 0
'        .FTP.RemoteInfo4FTP = ""
'        .FTP.UserName = ""
        
        .JSFileName = "menu"
        .GenDynAPI = False
        .CompileIECode = True
        .CompileNSCode = False
        .CompilehRefFile = False
        .SEOTweak = True
        
        ReDim .SecondaryProjects(0)
        
        .MenusOffset.RootMenusX = 0
        .MenusOffset.RootMenusY = 0
        .MenusOffset.SubMenusX = 0
        .MenusOffset.SubMenusY = 0
        
        .UnfoldingSound.onmouseover = ""
        
        .FontSubstitutions = ""
        
        .DoFormsTweak = False
        .DWSupport = False
        .NS4ClipBug = False
        '.OPHelperFunctions = False
        .ImageReadySupport = False
        .LotusDominoSupport = False
        
        .HideDelay = 200
        .SubMenusDelay = 150
        .RootMenusDelay = 15
        .SelChangeDelay = 0
        .AnimSpeed = 35
        
        .DXFilter = ""
        .BlinkEffect = 0
        .BlinkSpeed = 50
        
        .ExportHTMLParams = GenExpHTMLPref("", Project.Name, Project.FileName)
        
        .CustomOffsets = ""
        
        With .AutoScroll
            .maxHeight = 0
            .nColor = &H808080
            .hColor = &H202080
            .DnImage.NormalImage = vbNullString
            .DnImage.HoverImage = vbNullString
            .DnImage.w = 0
            .DnImage.h = 0
            .UpImage.NormalImage = vbNullString
            .UpImage.HoverImage = vbNullString
            .margin = 4
            .onmouseover = True
            .maxHeight = 0
        End With
    End With
    ReDim CustomSets(0)
    
    IsUpdating = True
    
    InitDir = ""
    
    LastSelected = ""
    LastItemFP = ""
    
    tvMenus.Nodes.Clear
    txtCaption.Text = ""
    tsCmdType.Tabs("tsOver").Selected = True
    cmbActionType.ListIndex = 0
    cmbTargetFrame.Visible = False
    cmbTargetMenu.Clear
    txtURL.Text = ""
    txtStatus.Text = ""
    mnuToolsDynAPI.Checked = False
    
    sbDummy.Panels(1).Text = ""
    
    IsUpdating = False
    
    If Not IsState Then
        UpdateTitleBar
        ResetUndoStates
    End If
    
    Erase FramesInfo.Frames
    Erase MenuGrps: ReDim MenuGrps(0)
    Erase MenuCmds: ReDim MenuCmds(0)
    
    If InMapMode Then
        RefreshMap
        tvMapView.Nodes(1).Selected = True
    End If
    InitLivePreview
    UpdateControls
    
    Project.HasChanged = False
    NewMenu = True

    On Error GoTo 0
    Exit Function

NewMenu_Error:

    MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Form.frmMain.NewMenu"

End Function

Private Sub ResetUndoStates()

    ReDim UndoStates(0)
    CurState = -1
    SaveState ""
    SynchUndoButtons

End Sub

Private Function ShowOpenDialog() As Boolean

    With cDlg
        .DialogTitle = GetLocalizedStr(255)
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNNoReadOnlyReturn
        .filter = GetLocalizedStr(256) + "|*.dmb"
        .FileName = ""
        Err.Clear
        On Error Resume Next
        .ShowOpen
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then
            ShowOpenDialog = False
            Exit Function
        End If
        Project.FileName = .FileName
    End With
    
    ShowOpenDialog = True

End Function

Private Sub LockFile(ByVal FileName As String)

    Dim ff As Integer
    
    On Error Resume Next
    
    If LenB(FileName) = 0 Then Exit Sub
    If Not Preferences.LockProjects Then Exit Sub
    
    FileName = Long2Short(FileName) + ".lck"
    
    ff = FreeFile
    Open FileName For Output As #ff
        Print #ff, App.FileDescription + " " + GetCurProjectVersion
        Print #ff, AppPath
    Close #ff
    
    SetAttr FileName, vbHidden

End Sub

Private Sub UnLockFile(ByVal FileName As String)

    On Error Resume Next

    Err.Clear
    If LenB(FileName) <> 0 Then
        FileName = Long2Short(FileName) + ".lck"
        If FileExists(FileName) Then
            SetAttr FileName, vbNormal
            Kill FileName
            If Preferences.LockProjects Then
                If Err.number <> 0 Then MsgBox GetLocalizedStr(844) + " " + vbCrLf + FileName, vbCritical + vbOKOnly, GetLocalizedStr(845)
            End If
        End If
    End If

End Sub

Private Function FileIsLocked(ByVal FileName As String, Optional lhIns As String) As Boolean

    On Error Resume Next
    
    If Preferences.LockProjects Then
        If LenB(FileName) <> 0 Then
            FileName = Long2Short(FileName) + ".lck"
            FileIsLocked = FileExists(FileName)
            If FileIsLocked Then
                lhIns = LoadFile(FileName)
            End If
        End If
    Else
        FileIsLocked = False
    End If
    
End Function

Friend Function LoadMenu(Optional ByVal File As String, Optional IsFromRecentList As Boolean, Optional IsState As Boolean, Optional IsPreset As Boolean) As Boolean

    Dim sStr As String
    Dim nLines As Integer
    Dim cLine As Integer
    Dim lhIns As String
    Dim i As Integer
        
    On Error GoTo chkError
    
    If Not NewMenu(IsState) Then Exit Function
    
    If IsState Then
        File = UndoStates(CurState).FileName
    Else
        If LenB(File) = 0 Then
            If Not ShowOpenDialog Then Exit Function
            File = Project.FileName
        Else
            If Not RemoveReadOnly(File) Then
                MsgBox GetLocalizedStr(257), vbInformation + vbOKOnly, GetLocalizedStr(543)
                LoadMenu = False
                Exit Function
            End If
        End If
        Project.HasChanged = False
    End If
    
    If FileIsLocked(File, lhIns) Then
        If MsgBox(File + " " + GetLocalizedStr(846) + vbCrLf + vbCrLf + lhIns + vbCrLf + GetLocalizedStr(847), vbQuestion + vbYesNo, GetLocalizedStr(543)) = vbYes Then
            UnLockFile File
        Else
            LoadMenu = False
            Exit Function
        End If
    Else
        LockFile File
    End If
    
    IsLoadingProject = True
    tvMenus.Visible = False
    Me.Enabled = False
    
    ff = FreeFile
    Open File For Binary As ff
        Do Until (LOF(ff) = Loc(ff)) Or sStr = "[RSC]"
            Line Input #ff, sStr
            If LenB(sStr) <> 0 Then nLines = nLines + 1
        Loop
    Close ff
    nLines = nLines - 2
    FloodPanel.caption = GetLocalizedStr(258)
    sbDummy.Style = sbrSimple
    
    Project = GetProjectProperties(File)
    #If DEVVER = 0 Then
        Project.GenDynAPI = False
        #If LITE = 1 Then
            For i = 0 To UBound(Project.UserConfigs)
                If Project.UserConfigs(i).Frames.UseFrames Then frmMain.ShowLITELImitationInfo 4
                Project.UserConfigs(i).Frames.UseFrames = False
            Next i
        #End If
    #Else
        If Project.RemoveImageAutoPosCode And Project.GenDynAPI Then
            Project.RemoveImageAutoPosCode = False
        End If
    #End If
    mnuToolsDynAPI.Checked = Project.GenDynAPI
    
    cmbTargetFrame.Clear
    If Project.UserConfigs(Project.DefaultConfig).Frames.UseFrames Then
        FramesInfo.FileName = Project.UserConfigs(Project.DefaultConfig).Frames.FramesFile
        GetFramesInfo
    End If
    
    UpdateTitleBar
    UpdateTargetFrameCombo
    
    LockWindowUpdate tvMenus.hwnd
    
    If (LOF(ff) = Loc(ff)) Then GoTo ExitSub
    Line Input #ff, sStr
    Do Until (LOF(ff) = Loc(ff)) Or sStr = "[RSC]"
        If LenB(sStr) <> 0 Then AddMenuGroup Mid$(sStr, 4)
        Do Until (LOF(ff) = Loc(ff)) Or sStr = "[RSC]"
            Line Input #ff, sStr
            If LenB(sStr) <> 0 Then
                cLine = cLine + 1: FloodPanel.Value = (cLine / nLines) * 100
                If Left$(sStr, 3) = "[C]" Then
                    AddMenuCommand Mid$(sStr, 6), True
                    If NagScreenIsVisible Then DoEvents
                Else
                    Exit Do
                End If
            End If
        Loop
        If (LOF(ff) = Loc(ff)) Then Exit Do
    Loop
    
    UpgradeProject GetCurProjectVersion
    
    LoadMenu = True
ExitSub:
    Close ff
    cSep = Chr(255) + Chr(255)
    
    LockWindowUpdate 0
    
    IsLoadingProject = False
    
    #If DEVVER = 0 Then
        Dim g As Integer
        For g = 1 To UBound(MenuGrps)
            If MenuGrps(g).IsTemplate Then
                MsgBox GetLocalizedStr(848), vbCritical + vbOKOnly, GetLocalizedStr(543)
                NewMenu
                Exit Function
            End If
        Next g
    #End If
    
    If Not IsState Then
        UpdateItemsLinks
        If tvMenus.Nodes.Count > 0 Then
            With tvMenus.Nodes(1)
                .Selected = True
                .EnsureVisible
            End With
        End If
        UpdateLivePreview
        DoEvents
        If InMapMode Then
            tvMapView.Nodes(1).Selected = True
        Else
            If tvMenus.Nodes.Count > 0 Then
                SelectItem tvMenus.Nodes(1)
            End If
        End If
        UpdateControls
        RefreshMap True
        
        If Not IsFromRecentList And Not IsPreset Then AddNewRecent
        ResetUndoStates
    End If
    
    FloodPanel.Value = 0
    sbDummy.Style = sbrNormal
    tvMenus.Visible = True
    Me.Enabled = True
    
    If (Preferences.VerifyLinksOptions.VerifyOptions And lvcVerifyWhenOpening) = lvcVerifyWhenOpening Then
        LinkVerifyMode = spmcAuto
        BrokenLinksReport
    End If
   
    Exit Function
    
chkError:
    Dim errNum As Long
    Dim errDes As String
    
    errNum = Err.number
    errDes = Err.Description
    If errNum = 53 Or errNum = 76 Then
        MsgBox GetLocalizedStr(261) + " " + File + " " + GetLocalizedStr(262), vbCritical + vbOKOnly, GetLocalizedStr(543)
    Else
        MsgBox GetLocalizedStr(263) + vbCrLf + GetLocalizedStr(544) + " (" & errNum & ") " + errDes, vbCritical + vbOKOnly, GetLocalizedStr(543)
    End If
    Project.HasChanged = False
    GoTo ExitSub

End Function

Private Sub UpgradeProject(dmbVer As String)

    Dim g As Integer
    Dim c As Integer
    Dim t As Integer
    Dim j As Integer
    Dim cVer As Long
    
    On Error Resume Next
    
    cVer = CLng(Project.version)
    
    #If LITE = 1 Then
        ApplyLITELimitations
    #End If
    
    If cVer < dmbVer Then
        FloodPanel.caption = GetLocalizedStr(849)
    Else
        Exit Sub
    End If

    If cVer < 302000 Then
        Project.DoFormsTweak = True
    End If
    
    If cVer < 301000 Then
        t = UBound(MenuGrps)
        For g = 1 To t
            FloodPanel.Value = g / t * 100
            With MenuGrps(g)
                .Corners.bottomCorner = .Corners.rightCorner
                .Corners.topCorner = .Corners.leftCorner
                .AlignmentStyle = ascVertical
            End With
        Next g
    End If
    
    If cVer < CLng(dmbVer) Then
        If cVer < 300000 Then
            FixOldProjectCommands
            FixOldProjectGroups
            Project.HasChanged = True
        End If
    End If
    
    FloodPanel.caption = GetLocalizedStr(849)
    
    If cVer < 300650 Then
        Project.CompileIECode = True
        Project.CompileNSCode = True
        Project.CompilehRefFile = True
    End If
    
    If cVer < 302022 Then
        t = UBound(Project.UserConfigs)
        For c = 0 To t
            FloodPanel.Value = c / t * 100
            Project.UserConfigs(c).OptmizePaths = (Project.UserConfigs(c).Type = ctcRemote)
        Next c
    End If
    
    If cVer < 305000 Then
        t = UBound(MenuCmds) + UBound(MenuGrps)
        For c = 0 To UBound(MenuCmds)
            FloodPanel.Value = c / t * 100
            With MenuCmds(c)
                .Sound.onclick = ""
                .Sound.onmouseover = ""
                .Actions.onmouseover.TargetMenuAlignment = IIf(MenuGrps(.parent).AlignmentStyle = ascHorizontal, gacBottomLeft, gacRightTop)
                .Actions.onclick.TargetMenuAlignment = IIf(MenuGrps(.parent).AlignmentStyle = ascHorizontal, gacBottomLeft, gacRightTop)
                .Actions.OnDoubleClick.TargetMenuAlignment = IIf(MenuGrps(.parent).AlignmentStyle = ascHorizontal, gacBottomLeft, gacRightTop)
                
                If Left(.Actions.onclick.TargetFrame, 4) <> "top." Then
                    .Actions.onclick.TargetFrame = "_self"
                End If
                
                If LenB(.BackImage.NormalImage) <> 0 Or LenB(MenuGrps(.parent).Image) <> 0 Then
                    .nBackColor = -2
                    MenuGrps(.parent).CmdsFXnColor = -2
                End If
            End With
        Next c
        For g = 0 To UBound(MenuGrps)
            FloodPanel.Value = (g + c - 1) / t * 100
            With MenuGrps(g)
                .Sound.onclick = ""
                .Sound.onmouseover = ""
                If Left(.Actions.onclick.TargetFrame, 4) <> "top." Then
                    .Actions.onclick.TargetFrame = "_self"
                End If
            End With
        Next g
        With Project.ToolBar
            .Width = 0
            .Height = 0
            .BorderColor = .BackColor
            .ContentsMarginH = .bOrder
            .ContentsMarginV = .bOrder
        End With
    End If
    
    If Project.AddIn.Name = "Dreamweaver Rollover Support" Then
        DisplayTip "Discontinued AddIns", "The 'Dreamweaver Rollover Support' AddIn is no longer required with this version of DHTML Menu Builder." + vbCrLf + vbCrLf + "A new option has been added to the File->Project Properties->Global Settings dialog to enable or disable this support." + vbCrLf + "Since your project was using this AddIn, DHTML Menu Builder has enabled this option for you.", False
        Project.AddIn.Name = ""
        Project.DWSupport = True
    End If
    If Project.AddIn.Name = "Navigator 4 CLIP bug" Then
        DisplayTip "Discontinued AddIns", "The 'Navigator 4 CLIP Bug' AddIn is no longer required with this version of DHTML Menu Builder." + vbCrLf + vbCrLf + "A new option has been added to the File->Project Properties->Global Settings dialog to enable or disable this support." + vbCrLf + "Since your project was using this AddIn, DHTML Menu Builder has enabled this option for you.", False
        Project.AddIn.Name = ""
        Project.NS4ClipBug = True
    End If
    
    If Project.AnimSpeed = 0 Then Project.AnimSpeed = 20
    If Project.HideDelay = 0 Then Project.HideDelay = 2000
    If Project.SubMenusDelay = 0 Then Project.SubMenusDelay = 200
    If Project.RootMenusDelay = 0 Then Project.RootMenusDelay = 15
    
    If Project.UserConfigs(0).Type = ctcLocal Then
        Project.UserConfigs(0).Type = ctcCDROM
    End If
    
    Project.UnfoldingSound.onclick = ""
    Project.UnfoldingSound.onmouseover = ""
    
    If cVer < 360004 Then
        Project.DXFilter = ""
    End If
    
    If cVer < 400000 Then
        If Project.ToolBar.CreateToolbar Then
            ReDim Project.Toolbars(1)
            Project.ToolBar.Name = GetLocalizedStr(400)
            Project.Toolbars(1) = Project.ToolBar
            ReDim Project.Toolbars(1).Groups(0)
        End If
        t = UBound(MenuGrps) + UBound(MenuCmds)
        For g = 1 To UBound(MenuGrps)
            FloodPanel.Value = g / t * 100
            With MenuGrps(g)
                If .IncludeInToolbar And Project.ToolBar.CreateToolbar Then
                    ReDim Preserve Project.Toolbars(1).Groups(UBound(Project.Toolbars(1).Groups) + 1)
                    Project.Toolbars(1).Groups(UBound(Project.Toolbars(1).Groups)) = .Name
                End If
                
                If .iCursor.CFile = "0" Then .iCursor.CFile = ""
            End With
        Next g
        
        For c = 1 To UBound(MenuCmds)
            FloodPanel.Value = (c + g - 1) / t * 100
            With MenuCmds(c)
                If .iCursor.CFile = "0" Then .iCursor.CFile = ""
                If .SeparatorPercent = 0 Then .SeparatorPercent = 80
            End With
        Next c
        
        FixLinks
    End If
    
    If cVer < 400000 Then
        ReDim Project.SecondaryProjects(0)
    End If
    
    If cVer < 402000 Then
        t = UBound(MenuGrps) + UBound(MenuCmds) + UBound(Project.Toolbars)
        For g = 1 To UBound(MenuGrps)
            FloodPanel.Value = g / t * 100
            With MenuGrps(g)
                If .DropShadowColor = 0 Then
                    .DropShadowSize = 0
                    .DropShadowColor = &H999999
                Else
                    .DropShadowSize = 3
                    .DropShadowColor = GenFilterColor(.DropShadowColor)
                End If
                .Transparency = .Transparency * 10
                
                With .Actions
                    If .onclick.Type = atcCascade And .onmouseover.Type = atcCascade Then
                        If .onclick.TargetMenu = .onmouseover.TargetMenu Then
                            .onclick.Type = atcNone
                        End If
                    End If
                End With
            End With
        Next g
        For c = 1 To UBound(MenuCmds)
            FloodPanel.Value = (g + c - 1) / t * 100
            With MenuCmds(c).Actions
                If .onclick.Type = atcCascade And .onmouseover.Type = atcCascade Then
                    If .onclick.TargetMenu = .onmouseover.TargetMenu Then
                        .onclick.Type = atcNone
                    End If
                End If
            End With
        Next c
        For j = 1 To UBound(Project.Toolbars)
            FloodPanel.Value = (c + g + j - 1) / t * 100
            With Project.Toolbars(j)
                .DropShadowColor = &H999999
                .DropShadowSize = 0
                .Transparency = 0
            End With
        Next j
        
        With Project
            If LenB(.NodeExpStatus) = 0 Then
                .NodeExpStatus = "+" + GetLocalizedStr(795) + "+"
            Else
                .NodeExpStatus = Replace(.NodeExpStatus, "+0+", "+")
                .NodeExpStatus = Replace(.NodeExpStatus, "+1+", "+")
            End If
        End With
    End If
    
    If cVer < 403005 Then
        Project.RootMenusDelay = 15
    End If
    
    If cVer < 403010 Then
        Project.LotusDominoSupport = False
    End If
    
    If cVer < 405001 Then
        Project.HideDelay = Project.HideDelay / 10
        If Project.HideDelay > 2000 Then Project.HideDelay = 2000
        Project.RemoveImageAutoPosCode = False
    End If
    
    If cVer < 405004 Then
        With Project.ExportHTMLParams
            .IncludeExpCol = False
            .ColAllStr = "Collapse All"
            .ExpAllStr = "Expand All"
            .ExpColPlacement = ecpcBottom
        End With
    End If
    
    If cVer < 406000 Then
        Project.AutoSelFunction = False
    End If
    
    If cVer < 407000 Then
        Project.KeyboardSupport = False
    End If
    
'    For g = 1 To UBound(MenuGrps)
'        With MenuGrps(g)
'            .WinStatus = ""
'            If .Actions.onclick.Type = atcURL And .Actions.onclick.url = "" Then .Actions.onclick.Type = atcNone
'        End With
'    Next g
'    For c = 1 To UBound(MenuCmds)
'        With MenuCmds(c)
'            .WinStatus = ""
'            If .Actions.onclick.Type = atcURL And .Actions.onclick.url = "" Then .Actions.onclick.Type = atcNone
'        End With
'    Next c

    If cVer < 409000 Then
        For g = 1 To UBound(MenuGrps)
            For c = 1 To UBound(MenuCmds)
                If MenuCmds(c).parent = g Then
                    With MenuCmds(c)
                        .CmdsFXhColor = MenuGrps(g).CmdsFXhColor
                        .CmdsFXnColor = MenuGrps(g).CmdsFXnColor
                        .CmdsFXNormal = MenuGrps(g).CmdsFXNormal
                        .CmdsFXOver = MenuGrps(g).CmdsFXOver
                        .CmdsFXSize = MenuGrps(g).CmdsFXSize
                        .CmdsMarginX = MenuGrps(g).CmdsMarginX
                        .CmdsMarginY = MenuGrps(g).CmdsMarginY
                    End With
                End If
            Next c
        Next g
    End If
    
    If cVer < 409002 Then
        Project.BlinkEffect = 0
        Project.BlinkSpeed = 50
    End If
    
    If cVer < 409007 Then
        With Project.AutoScroll
            .maxHeight = 0
            .nColor = &H808080
            .hColor = &H202080
            .DnImage.NormalImage = vbNullString
            .DnImage.HoverImage = vbNullString
            .DnImage.w = 0
            .DnImage.h = 0
            .UpImage.NormalImage = vbNullString
            .UpImage.HoverImage = vbNullString
            .margin = 4
            .onmouseover = True
            .maxHeight = 0
        End With
        For g = 1 To UBound(MenuGrps)
            With MenuGrps(g).scrolling
                If .maxHeight = 0 Then
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
                End If
            End With
        Next g
    End If
    
    If cVer < 409012 Then
        Project.SelChangeDelay = 0
        If Project.AddIn.Name <> "" Then
            Dim cCode As String
            cCode = LoadFile(AppPath + "addins/" + Project.AddIn.Name + ".ext")
            If InStr(cCode, "var lmcHS") = 0 Then
                cCode = Replace(cCode, "var lsc = null;", "var lsc = null;" + vbCrLf + vbTab + "var lmcHS = null;")
                SaveFile AppPath + "addins/" + Project.AddIn.Name + ".ext", cCode
            End If
        End If
    End If
    
    If cVer < 409016 Then
        Project.UseGZIP = False
    End If
    
    If cVer < 409018 Then
        For t = 1 To UBound(Project.Toolbars)
            Project.Toolbars(t).Compile = True
        Next t
        For c = 1 To UBound(MenuCmds)
            MenuCmds(c).Compile = True
        Next c
        For g = 1 To UBound(MenuGrps)
            MenuGrps(g).Compile = True
        Next g
    End If
    
    If cVer < 410000 Then
        With TipsSys
            .CanDisable = False
            .DialogTitle = "Upgrade Warning"
            .TipTitle = "Upgrading from previous versions"
            .Tip = "The project you're about to open was created with a previous version of DHTML Menu Builder." + vbCrLf + _
                    "While most projects can be automatically upgraded some projects may present problems, especially those using AddIns and Custom Offsets." + vbCrLf + _
                    "Please contact support (https://xfx.net/support) in case that you experience problems with your project."
            .Show
        End With
        
        c = 0
        For t = 1 To UBound(Project.Toolbars)
            If Project.Toolbars(t).Alignment = tbacAttached Then
                c = 1
                Exit For
            End If
        Next t
        If c = 1 Then
            With TipsSys
                .CanDisable = False
                .DialogTitle = "Upgrade Warning"
                .TipTitle = "Upgrading from previous versions"
                .Tip = "One or more of the toolbars in this project have been configured to attach itself to an image for positioning purposes." + vbCrLf + _
                        "Although this alignment method is still supported in this version it is highly recommended that you consider changing it to the new 'Free Flow' alignment option." + vbCrLf + _
                        "You can obtain more information about this new alignment option by clicking Help->Tutorials->Positioning the menus"
                .CanDisable = True
                .Show
            End With
        End If
        
        If Project.AddIn.Name = "RollerEffect" Then
            With TipsSys
                .CanDisable = False
                .DialogTitle = "Upgrade Warning"
                .TipTitle = "Upgrading from previous versions"
                .Tip = "This project is using the RollerEffect AddIn but this AddIn is no longer supported." + vbCrLf + _
                        "Your project will be automatically upgraded to use the new RollerEffect2 AddIn."
                .CanDisable = True
                .Show
            End With
            Project.AddIn.Name = "RollerEffect2"
        End If
    End If
    
    If cVer < 415001 Then
        For c = 1 To UBound(MenuCmds)
            MenuCmds(c).BackImage.AllowCrop = True
            MenuCmds(c).BackImage.Tile = True
        Next c
        For g = 1 To UBound(MenuGrps)
            MenuGrps(g).tbiBackImage.AllowCrop = True
            MenuGrps(g).tbiBackImage.Tile = True
        Next g
    End If
    
    If cVer < 420032 Then
        Project.CompileNSCode = False
        Project.CompilehRefFile = False
        Project.DoFormsTweak = False
        Project.LotusDominoSupport = False
        Project.DWSupport = False
        Project.SEOTweak = True
    End If
        
    Project.version = dmbVer
    
End Sub

Private Function GenFilterColor(Value As Long) As String

    Dim v As String
    
    v = Format$(Hex(255 - (Round(Value / 100 * 255))), "00")
    GenFilterColor = Val("&h" + v + v + v)

End Function

Private Sub FixLinks()

    Dim g As Integer
    Dim c As Integer
    
    Dim LocalRootWeb As String
    Dim RealRootWeb As String
    
    LocalRootWeb = GetRealLocal.RootWeb
    RealRootWeb = Project.UserConfigs(Project.DefaultConfig).RootWeb

    For g = 1 To UBound(MenuGrps)
        With MenuGrps(g)
            With .Actions.onclick
                If .Type = atcURL Or .Type = atcNewWindow Then
                    .url = FixLink(.url, LocalRootWeb, RealRootWeb)
                End If
            End With
        End With
    Next g
    
    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c)
            With .Actions.onclick
                If .Type = atcURL Or .Type = atcNewWindow Then
                    .url = FixLink(.url, LocalRootWeb, RealRootWeb)
                End If
            End With
        End With
    Next c

End Sub

Private Function FixLink(ByVal url As String, LocalRootWeb As String, RealRootWeb As String) As String

    Dim s As String
    Dim p As Integer

    If Not UsesProtocol(url) And Not IsExternalLink(url) Then
        If FileExists(LocalRootWeb + url) Then
            url = RealRootWeb + url
            Select Case Project.UserConfigs(Project.DefaultConfig).Type
                Case ctcLocal, ctcCDROM
                    url = SetSlashDir(url, sdBack)
                    s = "\"
                Case ctcRemote
                    url = SetSlashDir(url, sdFwd)
                    s = "/"
            End Select
            p = InStr(3, url, s)
            If p > 0 Then url = Left(url, p) + Replace(Mid(url, p + 1), s + s, s)
        End If
    End If
    
    FixLink = url

End Function

Private Sub FixOldProjectCommands()

    Dim c As Integer
    Dim ca As String
    Dim tc As Integer
    
    FloodPanel.caption = GetLocalizedStr(264)
    
    tc = UBound(MenuCmds)
    For c = 1 To tc
        FloodPanel.Value = c / tc * 100
        With MenuCmds(c)
            If .Actions.onclick.Type = atcURL Then
                ca = .Actions.onclick.url
                If Not UsesProtocol(ca) Then
                    If Left(ca, 7) <> "http://" And Left(ca, 7) <> "ftp://" Then
                        ca = Project.UserConfigs(0).RootWeb + SetSlashDir(ca, sdBack)
                    End If
                End If
                .Actions.onclick.url = RemoveDoubleSlashes(ca)
            End If
            If .Actions.onmouseover.Type = atcURL Then
                ca = .Actions.onmouseover.url
                If Not UsesProtocol(ca) Then
                    If Left(ca, 7) <> "http://" And Left(ca, 7) <> "ftp://" Then
                        ca = Project.UserConfigs(0).RootWeb + SetSlashDir(ca, sdBack)
                    End If
                End If
                .Actions.onmouseover.url = RemoveDoubleSlashes(ca)
            End If
        End With
    Next c
    
End Sub

Private Sub FixOldProjectGroups()

    Dim g As Integer
    Dim tg As Integer
    
    FloodPanel.caption = GetLocalizedStr(264)
    
    tg = UBound(MenuGrps)
    For g = 1 To tg
        FloodPanel.Value = g / tg * 100
        With MenuGrps(g)
            .Actions.onmouseover.Type = atcCascade
            .Actions.onmouseover.TargetMenu = g
        End With
    Next g
    
    Project.UserConfigs(Project.DefaultConfig).ImagesPath = Project.UserConfigs(Project.DefaultConfig).CompiledPath
    
End Sub

Public Sub SaveState(ByVal Description As String)

    Dim AddNewState As Boolean
    
    If IsLoadingProject Or IsReplacing Then Exit Sub
    
    Description = Replace(Description, "&", "")

    If Preferences.EnableUndoRedo Then
        If CurState <> -1 Then
            If Description <> UndoStates(CurState).Description Then
                AddNewState = True
            End If
        Else
            AddNewState = True
        End If
        If AddNewState Then
            CurState = CurState + 1
            ReDim Preserve UndoStates(CurState)
        End If
        With UndoStates(CurState)
            .Description = Description
            .FileName = StatesPath + "state" & CurState & ".dus"
        End With
        Project.HasChanged = CurState > 0
        SaveMenu True
    Else
        RefreshMap
        Project.HasChanged = True
        CurState = 0
    End If
    
    SynchUndoButtons

End Sub

Private Sub DoUndo()

    Dim SelItm As Integer
    Dim ProjectFileName As String
    
    On Error Resume Next
    
    SelItm = tvMenus.SelectedItem.Index
    
    ProjectFileName = Project.FileName
    
    CurState = CurState - 1
    LoadMenu , , True
    SynchUndoButtons
    
    Project.FileName = ProjectFileName
    
    tvMenus.Nodes(SelItm).Selected = True
    tvMenus.Nodes(SelItm).EnsureVisible
    UpdateControls
    RefreshMap True
    
    If CurState = 0 Then Project.HasChanged = False

End Sub

Private Sub DoRedo()

    Dim SelItm As Integer
    Dim ProjectFileName As String
    
    On Error Resume Next
    
    SelItm = tvMenus.SelectedItem.Index
    
    ProjectFileName = Project.FileName

    CurState = CurState + 1
    LoadMenu , , True
    SynchUndoButtons
    
    Project.FileName = ProjectFileName
    
    tvMenus.Nodes(SelItm).Selected = True
    tvMenus.Nodes(SelItm).EnsureVisible
    UpdateControls
    RefreshMap True
    
    Project.HasChanged = True

End Sub

Private Sub SynchUndoButtons()

    Dim Desc As String

    Desc = UndoStates(CurState).Description
    tbMenu.Buttons("tbUndo").Enabled = CurState > 0
    tbMenu.Buttons("tbUndo").ToolTipText = GetLocalizedStr(265) + " " + Desc
    tbMenu.Buttons("tbRedo").Enabled = CurState < UBound(UndoStates)
    If UBound(UndoStates) > CurState Then
        Desc = UndoStates(CurState + 1).Description
    Else
        Desc = ""
    End If
    tbMenu.Buttons("tbRedo").ToolTipText = GetLocalizedStr(266) + " " + Desc
    
    mnuEditUndo.Enabled = tbMenu.Buttons("tbUndo").Enabled
    mnuEditRedo.Enabled = tbMenu.Buttons("tbRedo").Enabled

End Sub

Friend Sub SynchViews()

    Dim tNode As Node
    Dim sNode As Node
    
    On Error GoTo ExitSub

    If Not InMapMode Then Exit Sub
    
    Set sNode = tvMenus.SelectedItem
    If sNode Is Nothing Then Exit Sub
    For Each tNode In tvMapView.Nodes
        If tNode.tag = sNode.key Then
            With tNode
                tNode.Selected = True
                tNode.EnsureVisible
            End With
            Exit For
        End If
    Next tNode
    
ExitSub:

End Sub

Private Sub UpdateTargetFrameCombo()

    Dim i As Integer
    Dim nf As Integer
    
    On Error Resume Next
    With Project
        cmbTargetFrame.Clear
        If .UserConfigs(.DefaultConfig).Frames.UseFrames Then
            nf = UBound(FramesInfo.Frames)
            For i = 1 To nf
                cmbTargetFrame.AddItem FramesInfo.Frames(i).Name
            Next i
        End If
        cmbTargetFrame.AddItem "_self"
        cmbTargetFrame.AddItem "_top"
        cmbTargetFrame.AddItem "_blank"
        cmbTargetFrame.AddItem "_parent"
        ResizeComboList cmbTargetFrame
    End With

End Sub

Private Sub AddNewRecent()

    Dim i As Integer
    Dim j As Integer
    Dim ProjectFile As String
    Dim tmpRecent As RecentFile
    
    ProjectFile = Project.FileName
    
    For i = UBound(RecentFiles) To 1 Step -1
        If RecentFiles(i).Title = Project.Name And LCase$(RecentFiles(i).Path) = LCase$(ProjectFile) Then
            tmpRecent = RecentFiles(i)
            For j = i To 2 Step -1
                RecentFiles(j) = RecentFiles(j - 1)
            Next j
            RecentFiles(1) = tmpRecent
            UpdateRecentFilesMenu
            Exit Sub
        End If
    Next i
    
    For i = UBound(RecentFiles) To 2 Step -1
        RecentFiles(i) = RecentFiles(i - 1)
    Next i
    RecentFiles(1).Title = Project.Name
    RecentFiles(1).Path = ProjectFile
    
    UpdateRecentFilesMenu

End Sub

Private Sub SaveRecent()

    Dim i As Integer
    Dim j As Integer
    Dim SetNull As Boolean

    For i = 1 To UBound(RecentFiles)
        SetNull = (LenB(RecentFiles(i).Title) = 0)
        If SetNull Then
            'SaveSetting App.EXEName, "RecentFiles", "Recent" & i, "" & cSep & ""
        Else
            j = j + 1
            SaveSetting App.EXEName, "RecentFiles", "Recent" & j, RecentFiles(i).Title & cSep & RecentFiles(i).Path
        End If
    Next i

End Sub

Private Sub GetRecentFiles()

    Dim i As Integer
    Dim fn As String
    Dim FT As String
    Dim k As Integer
    Dim j As Integer
    Dim Ok2Add As Boolean
    
    On Error Resume Next
    
    ReDim RecentFiles(1 To 8)

    For i = 0 To UBound(RecentFiles)
        fn = Split(GetSetting(App.EXEName, "RecentFiles", "Recent" & i, ""), cSep)(1)
        If FileExists(fn) Then
            Ok2Add = True
            For j = 1 To k
                If RecentFiles(j).Path = fn Then
                    Ok2Add = False
                    Exit For
                End If
            Next j
            If Ok2Add Then
                FT = Split(GetSetting(App.EXEName, "RecentFiles", "Recent" & i, ""), cSep)(0)
                k = k + 1
                RecentFiles(k).Title = FT
                RecentFiles(k).Path = fn
            End If
        End If
    Next i
    
    mnuFileOpenRecentOP.Checked = Preferences.OpenLastProject
    
    UpdateRecentFilesMenu

End Sub

Private Sub UpdateRecentFilesMenu()

    Dim i As Integer
    Dim j As Integer
    
    On Error Resume Next
    
    For i = 1 To mnuFileOpenRecentR.Count
        Unload mnuFileOpenRecentR(i)
    Next i
    
    mnuFileOpenRecentR(0).caption = GetLocalizedStr(817)
    If LenB(RecentFiles(1).Title) <> 0 Then
        For i = 1 To UBound(RecentFiles)
            If LenB(RecentFiles(i).Title) <> 0 Then
                If j > 0 Then Load mnuFileOpenRecentR(j)
                With mnuFileOpenRecentR(j)
                    .Enabled = True
                    .caption = RecentFiles(i).Title + " [" & EllipseText(dummyText, RecentFiles(i).Path, DT_PATH_ELLIPSIS) & "]"
                    .tag = RecentFiles(j).Path
                End With
                j = j + 1
            End If
        Next i
    Else
        With mnuFileOpenRecentR(0)
            .caption = GetLocalizedStr(110)
            .Enabled = False
        End With
    End If
    
    If Not IsDebug Then xMenu.Initialize Me
    
End Sub

Friend Sub UpdateTitleBar()

    caption = Project.Name + " - DHTML Menu Builder"
    #If LITE = 1 Then
        caption = caption + " LITE"
    #End If
    If IsDebug Then caption = caption + " [" + GetLocalizedStr(545) + "]"

End Sub

Friend Sub SaveMenu(Optional IsState As Boolean)

    Dim OriginalFileName As String
    Dim SaveOK As Boolean
    
    If IsState Then
        OriginalFileName = Project.FileName
        Project.FileName = UndoStates(CurState).FileName
    End If
    
    If LenB(Project.FileName) = 0 Then
        On Error Resume Next
        With cDlg
            .DialogTitle = GetLocalizedStr(267)
            .FileName = Project.FileName
            .filter = GetLocalizedStr(256) + "|*.dmb"
            .FilterIndex = 1
            .DefaultExt = ".dmb"
            .Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
            Err.Clear
            .ShowSave
            If Err.number = cdlCancel And LenB(Project.FileName) = 0 Then Exit Sub
            Project.FileName = .FileName
        End With
        On Error GoTo 0
    End If
    
    FloodPanel.caption = GetLocalizedStr(268)
    
    If Not IsState Then
        If (Preferences.VerifyLinksOptions.VerifyOptions And lvcVerifyWhenSaving) = lvcVerifyWhenSaving Then
            LinkVerifyMode = spmcAuto
            BrokenLinksReport
        End If
    End If
    
    SaveOK = SaveProject(IsState)
    
    If IsState Then
        If SaveOK Then
            If CurState = 0 Then
                If LenB(OriginalFileName) <> 0 Then SaveImages OriginalFileName
            Else
                SaveImages StatesPath + "state" & CurState - 1 & ".dus"
            End If
        End If
        Project.FileName = OriginalFileName
    Else
        If SaveOK Then
            SaveImages StatesPath + "state" & CurState & ".dus"
            Project.HasChanged = False
            ResetUndoStates
        End If
    End If
    
End Sub

Friend Sub FileSaveAs()

    On Error Resume Next
    
    With cDlg
        .CancelError = True
        .DialogTitle = GetLocalizedStr(269)
        .FileName = Project.FileName
        .Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
        .filter = GetLocalizedStr(256) + "|*.dmb"
        .FilterIndex = 1
        Err.Clear
        .ShowSave
        If Err.number = cdlCancel And LenB(Project.FileName) = 0 Then GoTo ExitSub
  
        Project.Name = GetFileName(.FileName)
        If GetFileExtension(Project.Name) = "dmb" Then Project.Name = Left(Project.Name, Len(Project.Name) - 4)
  
        If Preferences.EnableUndoRedo Then
            Project.FileName = .FileName
            SaveMenu
        Else
            FileCopy Project.FileName, .FileName
            Project.FileName = .FileName
        End If
        
        AddNewRecent
        UpdateTitleBar
    End With
    
    Exit Sub
    
ExitSub:
    Project.FileName = ""

End Sub

Private Sub DoUpgrade()

    If Project.HasChanged Then
        If Not NewMenu Then
            Exit Sub
        End If
    End If

    frmUpgrade.Show vbModal
    
    If Going2Upgrade Then
        RunShellExecute "Open", dlFileName, "/runDMB", TempPath, 0
        'Shell TempPath + dlFileName, vbMaximizedFocus
        Unload Me
    End If

End Sub

Private Sub MenuAddCommand()

    Dim sItem As Node
    Dim nid As Integer
    
    FinishRenaming
    IsRenaming = False
    
    Set sItem = tvMenus.SelectedItem
    If Not sItem Is Nothing Then
        If IsCommand(sItem.key) Or IsSeparator(sItem.key) Then Set sItem = sItem.parent
        If MenuGrps(GetID(sItem)).IsTemplate Then Exit Sub
    End If

    AddMenuCommand , , InMapMode
    
    If InMapMode Then
        nid = GetID
        MenuCmds(nid).Name = GetSecuenceName(False, "Command")
        tvMenus.SelectedItem.Text = MenuCmds(nid).Name
        LastItemFP = ""
        LastSelNode = ""
        RefreshMap
        DoEvents
        tvMapView.StartLabelEdit
    End If
    
    SaveState GetLocalizedStr(148) + " " + tvMenus.SelectedItem.Text
    
    If IsDebug Then SetupMenuMenu tvMenus.SelectedItem

End Sub

Private Sub MenuAddSeparator()

    Dim sItem As Node
    
    FinishRenaming
    IsRenaming = False
    
    KeepExpansions = False
    
    Set sItem = tvMenus.SelectedItem
    If Not sItem Is Nothing Then
        If IsCommand(sItem.key) Or IsSeparator(sItem.key) Then Set sItem = sItem.parent
        If MenuGrps(GetID(sItem)).IsTemplate Then Exit Sub
    End If

    AddCommandSeparator
    SaveState GetLocalizedStr(149)
    
    If InMapMode Then
        LastSelNode = ""
        LastItemFP = ""
        RefreshMap
    End If
    
    If IsDebug Then SetupMenuMenu tvMenus.SelectedItem

End Sub

Private Sub MenuAddGroup()

    Dim pNode As Node
    Dim tbIndex As Integer
    
    On Error Resume Next
    
    FinishRenaming
    IsRenaming = False
    
    IsTBMapSel = False
    KeepExpansions = False

    AddMenuGroup , InMapMode
    
    If InMapMode Then
        MenuGrps(UBound(MenuGrps)).Name = GetSecuenceName(True, "Group")
        Set pNode = tvMapView.SelectedItem
        If pNode Is Nothing Then
            tbIndex = IIf(CreateToolbar, 1, 0)
        Else
            Do Until pNode.parent Is Nothing
                Set pNode = pNode.parent
            Loop
            tbIndex = ToolbarIndexByName(pNode.tag)
        End If
        If tbIndex > 0 Then
            With Project.Toolbars(tbIndex)
                ReDim Preserve .Groups(UBound(.Groups) + 1)
                .Groups(UBound(.Groups)) = MenuGrps(UBound(MenuGrps)).Name
            End With
        End If
        tvMenus.SelectedItem.Text = MenuGrps(UBound(MenuGrps)).Name
        LastItemFP = ""
        LastSelNode = ""
        RefreshMap
        DoEvents
        tvMapView.StartLabelEdit
    End If
    
    SaveState GetLocalizedStr(147) + " " + tvMenus.SelectedItem.Text
    
    If IsDebug Then SetupMenuMenu tvMenus.SelectedItem

End Sub

Private Sub FillTargetGroupsCombo()

    Dim g As Integer

    cmbTargetMenu.Clear
    For g = 1 To UBound(MenuGrps)
        If InMapMode Then
            cmbTargetMenu.AddItem NiceGrpCaption(g)
        Else
            cmbTargetMenu.AddItem "[" + MenuGrps(g).Name + "]"
        End If
        cmbTargetMenu.ItemData(cmbTargetMenu.NewIndex) = g
    Next g
    ResizeComboList cmbTargetMenu

End Sub

Friend Sub UpdateControls()

    Dim IsG As Boolean
    Dim IsC As Boolean
    Dim IsS As Boolean
    Dim IsSG As Boolean
    Dim IsTpl As Boolean
    Dim IsRootWebValid As Boolean
    Dim sColor As Long
    
    On Error GoTo ForceCleanExit
    
    If IsUpdating Or IsLoadingProject Then Exit Sub
    
    IsUpdating = True
    IsRootWebValid = (LenB(GetRealLocal.RootWeb) <> 0)
    
    If tvMenus.SelectedItem Is Nothing Or IsTBMapSel Then
        IsC = False
        IsG = False
        IsS = False
        #If DEVVER = 1 Then
            If IsTBMapSel Then IsTpl = IsTemplate("")
        #End If
        chkCompile.Value = Abs(Project.Toolbars(ToolbarIndexByKey(tvMapView.SelectedItem.key)).Compile)
    Else
        IsG = IsGroup(tvMenus.SelectedItem.key)
        IsC = IsCommand(tvMenus.SelectedItem.key)
        IsS = IsSeparator(tvMenus.SelectedItem.key)
        If IsG Then IsSG = IsSubMenu(GetID)
        
        #If DEVVER = 1 Then
            If IsC Or IsS Then IsTpl = MenuGrps(MenuCmds(GetID).parent).IsTemplate
            If IsG Then IsTpl = MenuGrps(GetID).IsTemplate
        #End If
    End If
    #If DEVVER = 0 Then
        IsTpl = False
    #End If
        
    lblAlignment.Enabled = False
    icmbAlignment.Enabled = False
    txtURL.Text = ""
    
    If IsG Then
        tvMenus.SelectedItem.Image = GenGrpIcon(GetID)
        With MenuGrps(GetID)
            txtCaption.Text = xUNI2Unicode(.caption)
            chkEnabled.Value = Abs(Not .disabled)
            chkCompile.Value = Abs(.Compile)
            SelectAlignment .Alignment
            
            Select Case tsCmdType.SelectedItem.key
                Case "tsClick"
                    cmbActionType.ListIndex = .Actions.onclick.Type
                    Select Case .Actions.onclick.Type
                        Case atcURL, atcNewWindow
                            txtURL.Text = .Actions.onclick.url
                            If LenB(.Actions.onclick.TargetFrame) = 0 Then
                                cmbTargetFrame.Text = "_self"
                            Else
                                cmbTargetFrame.Text = .Actions.onclick.TargetFrame
                            End If
                        Case atcCascade
                            FillTargetGroupsCombo
                            On Error Resume Next
                            If .Actions.onclick.TargetMenu = 0 Then .Actions.onclick.TargetMenu = GetID
                            cmbTargetMenu.ListIndex = .Actions.onclick.TargetMenu - 1
                            On Error GoTo 0
                            lblAlignment.Enabled = True
                            icmbAlignment.Enabled = True
                    End Select
                Case "tsOver"
                    cmbActionType.ListIndex = .Actions.onmouseover.Type
                    Select Case .Actions.onmouseover.Type
                        Case atcURL, atcNewWindow
                            txtURL.Text = .Actions.onmouseover.url
                            If LenB(.Actions.onmouseover.TargetFrame) = 0 Then
                                cmbTargetFrame.Text = "_self"
                            Else
                                cmbTargetFrame.Text = .Actions.onclick.TargetFrame
                            End If
                        Case atcCascade
                            FillTargetGroupsCombo
                            On Error Resume Next
                            If .Actions.onmouseover.TargetMenu = 0 Then .Actions.onmouseover.TargetMenu = GetID
                            cmbTargetMenu.ListIndex = .Actions.onmouseover.TargetMenu - 1
                            On Error GoTo 0
                            lblAlignment.Enabled = True
                            icmbAlignment.Enabled = True
                    End Select
                Case "tsDoubleClick"
                    cmbActionType.ListIndex = .Actions.OnDoubleClick.Type
                    Select Case .Actions.OnDoubleClick.Type
                        Case atcURL, atcNewWindow
                            txtURL.Text = .Actions.OnDoubleClick.url
                            If LenB(.Actions.OnDoubleClick.TargetFrame) = 0 Then
                                cmbTargetFrame.Text = "_self"
                            Else
                                cmbTargetFrame.Text = .Actions.OnDoubleClick.TargetFrame
                            End If
                        Case atcCascade
                            FillTargetGroupsCombo
                            On Error Resume Next
                            If .Actions.OnDoubleClick.TargetMenu = 0 Then .Actions.OnDoubleClick.TargetMenu = GetID
                            cmbTargetMenu.ListIndex = .Actions.OnDoubleClick.TargetMenu - 1
                            On Error GoTo 0
                            lblAlignment.Enabled = True
                            icmbAlignment.Enabled = True
                    End Select
            End Select
            
            txtStatus.Text = xUNI2Unicode(.WinStatus)
            udLeading.Value = .Leading
            opAlignmentStyle(0).Value = .AlignmentStyle = ascVertical
            opAlignmentStyle(1).Value = .AlignmentStyle = ascHorizontal
        End With
        lblDataTitle.caption = GetLocalizedStr(124)
    ElseIf IsC Then
        tvMenus.SelectedItem.Image = GenCmdIcon(GetID)
        With MenuCmds(GetID)
            txtCaption.Text = xUNI2Unicode(.caption)
            txtURL.Text = .Actions.onclick.url
            txtStatus.Text = xUNI2Unicode(.WinStatus)
            chkEnabled.Value = Abs(Not .disabled)
            chkCompile.Value = Abs(.Compile)
            Select Case tsCmdType.SelectedItem.key
                Case "tsClick"
                    cmbActionType.ListIndex = .Actions.onclick.Type
                    Select Case .Actions.onclick.Type
                        Case atcURL, atcNewWindow
                            txtURL.Text = .Actions.onclick.url
                            If LenB(.Actions.onclick.TargetFrame) = 0 Then
                                cmbTargetFrame.Text = "_self"
                            Else
                                cmbTargetFrame.Text = .Actions.onclick.TargetFrame
                            End If
                        Case atcCascade
                            FillTargetGroupsCombo
                            If .Actions.onclick.TargetMenu = 0 Then
                                .Actions.onclick.TargetMenu = 1
                                DontRefreshMap = False
                                RefreshMap
                            End If
                            cmbTargetMenu.ListIndex = .Actions.onclick.TargetMenu - 1
                            lblAlignment.Enabled = True
                            icmbAlignment.Enabled = True
                        Case atcNone
                            txtURL.Text = ""
                    End Select
                    SelectAlignment .Actions.onclick.TargetMenuAlignment
                Case "tsOver"
                    cmbActionType.ListIndex = .Actions.onmouseover.Type
                    Select Case .Actions.onmouseover.Type
                        Case atcURL, atcNewWindow
                            txtURL.Text = .Actions.onmouseover.url
                            If LenB(.Actions.onclick.TargetFrame) = 0 Then
                                cmbTargetFrame.Text = "_self"
                            Else
                                cmbTargetFrame.Text = .Actions.onclick.TargetFrame
                            End If
                        Case atcCascade
                            FillTargetGroupsCombo
                            If .Actions.onmouseover.TargetMenu = 0 Then
                                .Actions.onmouseover.TargetMenu = 1
                                DontRefreshMap = False
                                RefreshMap
                            End If
                            cmbTargetMenu.ListIndex = .Actions.onmouseover.TargetMenu - 1
                            lblAlignment.Enabled = True
                            icmbAlignment.Enabled = True
                        Case atcNone
                            txtURL.Text = ""
                    End Select
                    SelectAlignment .Actions.onmouseover.TargetMenuAlignment
                Case "tsDoubleClick"
                    cmbActionType.ListIndex = .Actions.OnDoubleClick.Type
                    Select Case .Actions.OnDoubleClick.Type
                        Case atcURL, atcNewWindow
                            txtURL.Text = .Actions.OnDoubleClick.url
                            If LenB(.Actions.onclick.TargetFrame) = 0 Then
                                cmbTargetFrame.Text = "_self"
                            Else
                                cmbTargetFrame.Text = .Actions.onclick.TargetFrame
                            End If
                        Case atcCascade
                            FillTargetGroupsCombo
                            If .Actions.OnDoubleClick.TargetMenu = 0 Then
                                .Actions.OnDoubleClick.TargetMenu = 1
                                DontRefreshMap = False
                                RefreshMap
                            End If
                            cmbTargetMenu.ListIndex = .Actions.OnDoubleClick.TargetMenu - 1
                            lblAlignment.Enabled = True
                            icmbAlignment.Enabled = True
                        Case atcNone
                            txtURL.Text = ""
                    End Select
                    SelectAlignment .Actions.OnDoubleClick.TargetMenuAlignment
            End Select
        End With
        lblDataTitle.caption = GetLocalizedStr(102)
    ElseIf IsS Then
        chkCompile.Value = Abs(MenuCmds(GetID).Compile)
    End If
    
    txtCaption.Enabled = IsC Or IsG
    chkEnabled.Enabled = IsC Or IsG
    chkCompile.Enabled = True
    opAlignmentStyle(0).Enabled = IsG
    opAlignmentStyle(1).Enabled = IsG
    lblLayout.Enabled = IsG
    lblASVertical.Enabled = IsG
    lblASHorizontal.Enabled = IsG

    Select Case cmbActionType.ListIndex
        Case atcNone
            txtURL.Visible = True
            txtURL.Enabled = False
            cmdBrowse.Enabled = False
            cmdTargetFrame.Enabled = False
            cmdBookmark.Enabled = False
            lblActionName.Enabled = False
            cmdWinParams.Visible = False
            cmdFindTargetGroup.Visible = False
            
            cmbTargetMenu.Visible = False
            cmbTargetMenu.Enabled = False
        Case atcURL
            txtURL.Visible = True
            txtURL.Enabled = (IsC Or IsG) And -chkEnabled.Value
            cmdBrowse.Enabled = (IsC Or IsG) And IsRootWebValid And -chkEnabled.Value
            cmdWinParams.Visible = False
            cmdFindTargetGroup.Visible = False
            
            'cmbTargetFrame.Visible = (IsC Or IsG)
            'cmbTargetFrame.Enabled = cmbTargetFrame.Visible And -chkEnabled.Value
            
            cmdTargetFrame.Visible = (IsC Or IsG)
            cmdTargetFrame.Enabled = cmdTargetFrame.Visible And -chkEnabled.Value
            
            cmdBookmark.Visible = (IsC Or IsG)
            SetBookmarkState
            
            cmbTargetMenu.Visible = False
            
            lblActionName.Enabled = IsC Or IsG
            lblActionName.caption = GetLocalizedStr(125)
        Case atcCascade
            txtURL.Visible = False
            cmdWinParams.Visible = False
            cmdFindTargetGroup.Visible = True
            cmdFindTargetGroup.Enabled = (cmbTargetMenu.ListIndex >= 0)
            
            'cmbTargetFrame.Visible = False
            cmbTargetMenu.Visible = True
            cmbTargetMenu.Enabled = True
            
            lblActionName.Enabled = True
            lblActionName.caption = GetLocalizedStr(109)
        Case atcNewWindow
            txtURL.Visible = True
            txtURL.Enabled = IsC Or IsG
            cmdBrowse.Enabled = IsC Or IsG
            cmdWinParams.Visible = True
            cmdFindTargetGroup.Visible = False
            
            cmbTargetMenu.Visible = False
            cmdTargetFrame.Visible = False
            cmdBookmark.Visible = False
            
            lblActionName.Enabled = IsC Or IsG
            lblActionName.caption = "URL"
    End Select
    
    cmdWinParams.Enabled = (cmdWinParams.Enabled And Not IsS) Or (IsC Or IsG)
    cmbActionType.Enabled = (IsC Or IsG) And -chkEnabled.Value
    lblActionType.Enabled = (IsC Or IsG) And -chkEnabled.Value
    tsCmdType.Enabled = (IsC Or IsG) And -chkEnabled.Value
    lblActionName.Enabled = lblActionName.Enabled And ((IsC Or IsG) And -chkEnabled.Value)
    cmbTargetMenu.Enabled = (IsC Or IsG) And -chkEnabled.Value
    
    If IsG Then
        If IsSubMenu(GetID) And InMapMode Then
            If BelongsToToolbar(GetID, True) = 0 Then
                lblAlignment.Enabled = False
                icmbAlignment.Enabled = False
            End If
        End If
    End If
    
    lblAlignment.Enabled = lblAlignment.Enabled And -chkEnabled.Value
    icmbAlignment.Enabled = lblAlignment.Enabled And -chkEnabled.Value
    
    'txtStatus.Enabled = (-chkEnabled.Value And IsC Or (-chkEnabled.Value And IsG And CreateToolbar)) And Not -chkLockStatus.Value
    'chkLockStatus.Enabled = IsC Or (IsG And CreateToolbar)
    
    txtStatus.Enabled = (IsC Or IsG) And -chkEnabled.Value
    lblCaption.Enabled = txtCaption.Enabled
    lblStatus.Enabled = txtStatus.Enabled
    
    #If DEVVER = 1 Then
    If Not tvMenus.SelectedItem Is Nothing Then
        If IsTemplate(tvMenus.SelectedItem.key) Then
            txtCaption.Enabled = False
            chkEnabled.Enabled = False
            chkCompile.Enabled = False
            tsCmdType.Enabled = False
            txtStatus.Enabled = False
            icmbAlignment.Enabled = False
            opAlignmentStyle(0).Enabled = False
            opAlignmentStyle(1).Enabled = False
            cmbActionType.Enabled = False
            cmdTargetFrame.Enabled = False
            cmdBookmark.Enabled = False
            txtURL.Enabled = False
            cmdFindTargetGroup.Enabled = False
            cmbTargetMenu.Enabled = False
            cmdBrowse.Enabled = False
        End If
    End If
    #End If
    
    With tbCmd
        .Buttons("tbColor").Enabled = IsC Or IsG Or IsS
        .Buttons("tbFont").Enabled = IsC Or (IsG And Not IsSG)
        .Buttons("tbCursor").Enabled = IsC Or (IsG And Not IsSG)
        .Buttons("tbImage").Enabled = IsC Or IsG
        .Buttons("tbMargins").Enabled = IsG
        .Buttons("tbSelFX").Enabled = IsC Or (IsG And Not IsSG)
        .Buttons("tbFX").Enabled = IsG
        .Buttons("tbSound").Enabled = IsC Or IsG
        .Buttons("tbUp").Enabled = (tvMenus.Nodes.Count > 0)
        .Buttons("tbDown").Enabled = (tvMenus.Nodes.Count > 0)
        If InMapMode Then
            If IsTBMapSel Then
                If Left(tvMapView.SelectedItem.key, 4) = "TBK(" Then
                    .Buttons("tbUp").Enabled = False
                    .Buttons("tbDown").Enabled = False
                    chkCompile.Enabled = False
                End If
            End If
        End If
    End With
    
    udLeading.Enabled = IsG
    If IsG Then
        sColor = RGB(0, 0, 128)
    Else
        sColor = &H808080
    End If
    shpSpc(0).FillColor = sColor
    shpSpc(1).FillColor = sColor
    shpSpc(2).FillColor = sColor
    
    With tbMenu
        .Buttons("tbAddGrp").Enabled = Not IsTpl
        .Buttons("tbAddSubGrp").Visible = InMapMode
        If IsC Then
            .Buttons("tbAddSubGrp").Enabled = (MenuCmds(GetID).Actions.onmouseover.Type <> atcCascade)
        Else
            .Buttons("tbAddSubGrp").Enabled = False
        End If
        .Buttons("tbAddSubGrp").Enabled = .Buttons("tbAddSubGrp").Enabled And Not IsTpl
        .Buttons("tbAddCmd").Enabled = (tvMenus.Nodes.Count > 0) And Not IsTBMapSel And Not IsTpl
        .Buttons("tbAddSep").Enabled = (tvMenus.Nodes.Count > 0) And Not IsTBMapSel And Not IsTpl
        If IsTBMapSel Then
            .Buttons("tbCopy").Enabled = tvMapView.SelectedItem.key <> "TBK(No Toolbar)"
        Else
            .Buttons("tbCopy").Enabled = (IsG Or IsC Or IsS)
        End If
        .Buttons("tbPaste").Enabled = .Buttons("tbCopy").Enabled
        .Buttons("tbRemove").Enabled = (tvMenus.Nodes.Count) > 0 And Not IsTpl
        If InMapMode Then
            If IsTBMapSel Then
                .Buttons("tbRemove").Enabled = ToolbarIndexByName(tvMapView.SelectedItem.tag) > 0 And Not IsTpl
            End If
        End If

        mnuEditCopy.Enabled = .Buttons("tbCopy").Enabled
        mnuEditPaste.Enabled = .Buttons("tbPaste").Enabled
        mnuEditDelete.Enabled = .Buttons("tbRemove").Enabled
        mnuEditRename.Enabled = .Buttons("tbRemove").Enabled
    End With
    
    With tbMenu2
        .Buttons("tbCompile").Enabled = (UBound(MenuGrps) > 0)
        .Buttons("tbPublish").Enabled = .Buttons("tbCompile").Enabled
        .Buttons("tbPreview").Enabled = .Buttons("tbCompile").Enabled
        .Buttons("tbHotSpotEditor").Enabled = .Buttons("tbCompile").Enabled
        
        mnuToolsPreview.Enabled = .Buttons("tbPreview").Enabled
        mnuToolsCompile.Enabled = .Buttons("tbCompile").Enabled
        mnuToolsPublish.Enabled = .Buttons("tbPublish").Enabled
        mnuToolsHotSpotsEditor.Enabled = .Buttons("tbHotSpotEditor").Enabled
        mnuToolsInstallMenusAILC.Enabled = FileExists(GetRealLocal.HotSpotEditor.HotSpotsFile) Or CreateToolbar And mnuToolsCompile.Enabled
        mnuToolsInstallMenusAIFLC.Enabled = GetRealLocal.Frames.UseFrames
        mnuToolsInstallMenusAIRLC.Enabled = IsRootWebValid
        mnuToolsInstallMenus.Enabled = IsRootWebValid
        mnuToolsSecProj.Enabled = Not GetRealLocal.Frames.UseFrames
        mnuFileExportHTML.Enabled = mnuToolsCompile.Enabled
        mnuFileSaveAsPreset.Enabled = mnuToolsCompile.Enabled
        
        If Preferences.EnableUndoRedo Then
            mnuFileSaveAs.Enabled = True
        Else
            mnuFileSaveAs.Enabled = FileExists(Project.FileName)
        End If
    End With
    
    Dim SelObj As Object
    If IsC Then
        With MenuCmds(GetID).Actions
            tsCmdType.Tabs("tsClick").Image = IIf(.onclick.Type <> atcNone, ilTabs.ListImages("ClickON").Index, ilTabs.ListImages("ClickOFF").Index)
            tsCmdType.Tabs("tsDoubleClick").Image = IIf(.OnDoubleClick.Type <> atcNone, ilTabs.ListImages("DblClickON").Index, ilTabs.ListImages("DblClickOFF").Index)
            tsCmdType.Tabs("tsOver").Image = IIf(.onmouseover.Type <> atcNone, ilTabs.ListImages("MoveON").Index, ilTabs.ListImages("MoveOFF").Index)
        End With
    ElseIf IsG Then
        With MenuGrps(GetID).Actions
            tsCmdType.Tabs("tsClick").Image = IIf(.onclick.Type <> atcNone, ilTabs.ListImages("ClickON").Index, ilTabs.ListImages("ClickOFF").Index)
            tsCmdType.Tabs("tsDoubleClick").Image = IIf(.OnDoubleClick.Type <> atcNone, ilTabs.ListImages("DblClickON").Index, ilTabs.ListImages("DblClickOFF").Index)
            tsCmdType.Tabs("tsOver").Image = IIf(.onmouseover.Type <> atcNone, ilTabs.ListImages("MoveON").Index, ilTabs.ListImages("MoveOFF").Index)
        End With
    Else
        tsCmdType.Tabs("tsClick").Image = ilTabs.ListImages("ClickOFF").Index
        tsCmdType.Tabs("tsDoubleClick").Image = ilTabs.ListImages("DblClickOFF").Index
        tsCmdType.Tabs("tsOver").Image = ilTabs.ListImages("MoveOFF").Index
    End If
    
'    If IsFPAddIn Then
'        tbMenu.Buttons("tbOpen").Visible = False
'        tbMenu2.Buttons("tbCompile").Visible = False
'        tbMenu2.Buttons("tbHotSpotEditor").Visible = False
'        tbMenu2.Buttons("tbSep01").Visible = False
'        mnuToolsCompile.Visible = False
'        mnuToolsHotSpotsEditor.Visible = False
'        mnuToolsInstallMenus.Visible = False
'        mnuFileOpen.Visible = False
'        mnuFileOpenRecent.Visible = False
'        mnuToolsSep01.Visible = False
'    End If
    
ForceCleanExit:

    UpdateEasyControls

    DontRefreshMap = False

    UpdateStatusbar IsC, IsG, IsS
    
    If Preferences.UseLivePreview Then
        sbLP_GroupMode.Enabled = (IsG And Not IsSG)
        DoDelayedLivePreview
    End If
    
    IsUpdating = False

End Sub

Private Sub UpdateEasyControls()

    Dim IsG As Boolean
    Dim IsC As Boolean
    Dim IsS As Boolean
    Dim IsTpl As Boolean
    
    If Not Preferences.UseEasyActions Then Exit Sub
    
    If tvMenus.SelectedItem Is Nothing Or IsTBMapSel Then
        IsC = False
        IsG = False
        IsS = False
    Else
        IsG = IsGroup(tvMenus.SelectedItem.key)
        IsC = IsCommand(tvMenus.SelectedItem.key)
        IsS = IsSeparator(tvMenus.SelectedItem.key)
    End If

    sbEasyLink.Enabled = (IsC Or IsG)
    txtEasyLink.Enabled = (IsC Or IsG)
    sbEasyTargetFrame.Enabled = (IsC Or IsG)
    sbEasyBookmark.Enabled = (IsC Or IsG)
    chkEasyEnableNewWindow.Enabled = (IsC Or IsG)
    sbEasyNewWindow.Enabled = (IsC Or IsG)
    If chkEasyEnableNewWindow.Value = vbChecked Then
        sbEasyTargetFrame.Enabled = False
        sbEasyBookmark.Enabled = False
    End If

    If IsC Then
        Dim c As MenuCmd
        c = MenuCmds(GetID)
        
        txtEasyLink.Text = c.Actions.onclick.url
        sbEasyTargetFrame.ToolTipText = cmdTargetFrame.ToolTipText
        sbEasyBookmark.ToolTipText = cmdBookmark.ToolTipText
        chkEasyEnableNewWindow.Value = IIf(c.Actions.onclick.Type = atcNewWindow, vbChecked, vbUnchecked)
        
        cmbEasySubMenuAction.Clear
        If c.Actions.onmouseover.Type = atcNone Then
            cmbEasySubMenuAction.AddItem "(none)"
            cmbEasySubMenuAction.AddItem String(10, "-")
            cmbEasySubMenuAction.AddItem "Create SubMenu..."
            
            cmbEasySubMenuAction.ItemData(0) = 0
            cmbEasySubMenuAction.ItemData(1) = 0
            cmbEasySubMenuAction.ItemData(2) = 1
            
            cmbEasySubMenuAction.ListIndex = 0
            
            icmbEasyAlignment.Enabled = False
        Else
            cmbEasySubMenuAction.AddItem "Display On Mouse Over"
            cmbEasySubMenuAction.AddItem String(10, "-")
            cmbEasySubMenuAction.AddItem "Remove SubMenu..."
            
            cmbEasySubMenuAction.ItemData(0) = 0
            cmbEasySubMenuAction.ItemData(1) = 0
            cmbEasySubMenuAction.ItemData(2) = 2
            
            cmbEasySubMenuAction.ListIndex = 0
            
            icmbEasyAlignment.Enabled = True
        End If
    End If
    
    If IsG Then
        Dim g As MenuGrp
        g = MenuGrps(GetID)
        
        txtEasyLink.Text = g.Actions.onclick.url
        sbEasyTargetFrame.ToolTipText = cmdTargetFrame.ToolTipText
        sbEasyBookmark.ToolTipText = cmdBookmark.ToolTipText
        chkEasyEnableNewWindow.Value = IIf(g.Actions.onclick.Type = atcNewWindow, vbChecked, vbUnchecked)
        
        If IsSubMenu(GetID) Then
            txtEasyLink.Enabled = False
            sbEasyLink.Enabled = False
            sbEasyTargetFrame.Enabled = False
            sbEasyBookmark.Enabled = False
            chkEasyEnableNewWindow.Enabled = False
            icmbEasyAlignment.Enabled = False
            cmbEasySubMenuAction.Enabled = False
        Else
            cmbEasySubMenuAction.Enabled = True
            cmbEasySubMenuAction.Clear
            If g.Actions.onmouseover.Type = atcNone Then
                cmbEasySubMenuAction.AddItem "(none)"
                cmbEasySubMenuAction.AddItem String(10, "-")
                cmbEasySubMenuAction.AddItem "Create SubMenu..."
                
                cmbEasySubMenuAction.ItemData(0) = 0
                cmbEasySubMenuAction.ItemData(1) = 0
                cmbEasySubMenuAction.ItemData(2) = 1
                
                cmbEasySubMenuAction.ListIndex = 0
                
                icmbEasyAlignment.Enabled = False
            Else
                cmbEasySubMenuAction.AddItem "Display On Mouse Over"
                cmbEasySubMenuAction.AddItem String(10, "-")
                cmbEasySubMenuAction.AddItem "Remove SubMenu..."
                
                cmbEasySubMenuAction.ItemData(0) = 0
                cmbEasySubMenuAction.ItemData(1) = 0
                cmbEasySubMenuAction.ItemData(2) = 2
                
                cmbEasySubMenuAction.ListIndex = 0
                
                icmbEasyAlignment.Enabled = True
            End If
        End If
    End If
    
    With tsCmdType
        picEasy.Move .Left, .Top, .Width, .Height
        picEasy.ZOrder 0
        picEasy.BorderStyle = 0
        picEasyLinkBtns.BorderStyle = 0
    End With

End Sub

Private Sub cmbEasySubMenuAction_Click()

    Select Case cmbEasySubMenuAction.ItemData(cmbEasySubMenuAction.ListIndex)
        Case 1
            MenuAddSubGroup
        Case 2
            tsCmdType.Tabs(1).Selected = True
            tsCmdType_Click
    
            cmbActionType.ListIndex = 0
    End Select

End Sub

Private Sub ResizeEasyControls()

    If Not Preferences.UseEasyActions Then Exit Sub

    With tsCmdType
        picEasy.Move .Left, .Top, .Width, .Height
    End With
    
    With picEasy
        frmEasyLink.Width = .Width - 2 * frmEasyLink.Left
        frmEasySubMenu.Width = frmEasyLink.Width
        
        sbEasyLink.Left = .Width - sbEasyLink.Width - txtEasyLink.Left
        txtEasyLink.Width = sbEasyLink.Left - 2 * txtEasyLink.Left
        
        icmbEasyAlignment.Width = sbEasyLink.Left + sbEasyLink.Width - icmbEasyAlignment.Left
    End With
    
End Sub

Private Sub SetBookmarkState()

    Dim fn As String
    Dim miActions As ActionEvents

    cmdBookmark.Enabled = LenB(txtURL.Text) <> 0 And Not UsesProtocol(txtURL.Text) And txtURL.Enabled

    cmdBookmark.ToolTipText = "Bookmark"
    If cmdBookmark.Enabled And cmdBookmark.Visible Then
        fn = txtURL.Text
        If InStr(fn, "#") Then
            cmdBookmark.ToolTipText = cmdBookmark.ToolTipText + " (" + Mid(fn, InStrRev(fn, "#") + 1) + ")"
        End If
    End If
    
    fn = ""
    cmdTargetFrame.ToolTipText = GetLocalizedStr(851)
    If IsGroup(tvMenus.SelectedItem.key) Then
        miActions = MenuGrps(GetID).Actions
    Else
        miActions = MenuCmds(GetID).Actions
    End If
    With miActions
        Select Case tsCmdType.SelectedItem.key
            Case "tsOver"
                If .onmouseover.Type = atcURL Then fn = .onmouseover.TargetFrame
            Case "tsClick"
                If .onclick.Type = atcURL Then fn = .onclick.TargetFrame
            Case "tsDoubleClick"
                If .OnDoubleClick.Type = atcURL Then fn = .OnDoubleClick.TargetFrame
        End Select
    End With
    If LenB(fn) <> 0 Then cmdTargetFrame.ToolTipText = cmdTargetFrame.ToolTipText + " (" + fn + ")"
    
End Sub

Private Sub SelectAlignment(v As GroupAlignmentConstants)

    Dim i As Integer
    
    For i = 1 To icmbAlignment.ComboItems.Count
        If Val(icmbAlignment.ComboItems(i).tag) = v Then
            icmbAlignment.ComboItems(i).Selected = True
            ' Easy Combo
            icmbEasyAlignment.ComboItems(i).Selected = True
            Exit For
        End If
    Next i

End Sub

Private Sub UpdateStatusbar(IsC As Boolean, IsG As Boolean, IsS As Boolean)

    Dim pTB As Integer

    With sbDummy.Panels(1)
        If InMapMode And (Not IsC And Not IsG And Not IsS) Then
            If tvMapView.SelectedItem Is Nothing Then
                .Text = ""
            Else
                .Text = "Toolbar #" & ToolbarIndexByKey(tvMapView.SelectedItem.key) & ": " + tvMapView.SelectedItem.Text
            End If
        Else
            If IsC Then .Text = GetLocalizedStr(271) + ": " + tvMenus.SelectedItem.Text + IIf(LenB(txtCaption.Text) <> 0, " (" + txtCaption.Text + ")", "")
            If IsG Then
                pTB = MemberOf(GetID)
                .Text = GetLocalizedStr(270) + ": " + tvMenus.SelectedItem.Text + IIf(LenB(txtCaption.Text) <> 0, " (" + txtCaption.Text + ")", "") + IIf(CreateToolbar And pTB <> 0, " " + GetLocalizedStr(852) + " " + Project.Toolbars(pTB).Name, "")
            End If
            If IsS Then .Text = GetLocalizedStr(272)
        End If
        
        If InMapMode Then
            If Not tvMapView.SelectedItem Is Nothing Then
                If tvMapView.SelectedItem.ForeColor = Preferences.NoCompileItem Then
                    .Text = .Text + " [" + GetLocalizedStr(987) + "]"
                End If
            End If
        End If
    End With
    
    sbDummy.Panels(2).Text = GetLocalizedStr(571) + ": " + Project.UserConfigs(Project.DefaultConfig).Name

End Sub

Friend Sub UpdateLivePreview()

    If Not Preferences.UseLivePreview Or IsLoadingProject Or IsReplacing Then Exit Sub
    DoDelayedLivePreview

End Sub

Private Sub ShowColorDlg()

    If tvMenus.SelectedItem Is Nothing Then Exit Sub

    If IsGroup(tvMenus.SelectedItem.key) Then
        frmGrpColor.Show vbModal
    Else
        frmColor.Show vbModal
    End If
    
    UpdateLivePreview

End Sub

Private Sub ShowCursorDlg()

    If tvMenus.SelectedItem Is Nothing Then Exit Sub

    If IsCommand(tvMenus.SelectedItem.key) Then
        frmCursor.Show vbModal
    Else
        frmGrpCursor.Show vbModal
    End If
    
    UpdateLivePreview

End Sub

Private Sub ShowFontDlg()

    If tvMenus.SelectedItem Is Nothing Then Exit Sub

    If IsCommand(tvMenus.SelectedItem.key) Then
        frmFont.Show vbModal
    Else
        frmGrpFont.Show vbModal
    End If
    
    UpdateLivePreview

End Sub

Private Sub ShowSelFXDlg()

    If tvMenus.SelectedItem Is Nothing Then Exit Sub

    If IsGroup(tvMenus.SelectedItem.key) Then
        frmGrpSelectionFX.Show vbModal
    Else
        frmSelFX.Show vbModal
    End If
    
    UpdateLivePreview

End Sub

Private Sub ShowSFXDlg()

'    #If LITE = 1 Then
'        ShowLITELImitationInfo 2
'    #Else
'        frmGrpFX.Show vbModal
'    #End If
    
    frmGrpFX.Show vbModal
    UpdateLivePreview

End Sub

#If LITE = 1 Then
Friend Sub ShowLITELImitationInfo(id As Integer)

    Dim Msg As String

    Select Case id
        Case 1
            Msg = "This project cannot be compiled using the LITE version of DHTML Menu Builder." + vbCrLf + _
                    "The LITE version supports projects with just one toolbar."
        Case 2
            Msg = "This feature is not supported in the LITE version of DHTML Menu Builder."
        Case 3
            Msg = "The LITE version does not support the 'max' Code Optimization setting."
        Case 4
            Msg = "Frames support is not available in the LITE version of DHTML Menu Builder."
    End Select

    MsgBox Msg, vbInformation + vbOKOnly, "LITE version limitations"

End Sub

Private Sub ApplyLITELimitations()

    Dim i As Integer

    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            .DropShadowSize = 0
            .Transparency = 0
            '.fHeight = 0
            '.fWidth = 0
            .Compile = True
        End With
    Next i
    
    For i = 1 To UBound(MenuCmds)
        MenuCmds(i).Compile = True
    Next i
    
    With Project
        For i = 1 To UBound(.Toolbars)
            With .Toolbars(i)
                .DropShadowSize = 0
                .Transparency = 0
                .FollowHScroll = False
                .FollowVScroll = False
                '.Height = 0
                '.Width = 0
                .Condition = ""
                .Compile = True
            End With
        Next i
        
        For i = 1 To UBound(.UserConfigs)
            .UserConfigs(i).Frames.UseFrames = False
        Next i
        
        With .MenusOffset
            .RootMenusX = 0
            .RootMenusY = 0
            .SubMenusX = 0
            .SubMenusY = 0
        End With
        .CustomOffsets = ""
        .AddIn.Name = ""
        If .CodeOptimization = cocAggressive Then .CodeOptimization = cocNormal
        .UseGZIP = False
        .DXFilter = ""
        .KeyboardSupport = False
        ReDim .SecondaryProjects(0)
        .AutoScroll.maxHeight = 0
        .StatusTextDisplay = socStatus
        .AutoSelFunction = False
    End With
End Sub
#End If

Private Sub ShowImageDlg()

    If tvMenus.SelectedItem Is Nothing Then Exit Sub

    If IsGroup(tvMenus.SelectedItem.key) Then
        frmGrpImage.Show vbModal
    Else
        frmImage.Show vbModal
    End If
    
    UpdateLivePreview

End Sub

Friend Sub RemoveItem(Optional Force As Boolean = False, Optional RemoveTemplate As Boolean = False, Optional ForceRecurse As Boolean)

    Dim Ans As VbMsgBoxResult
    Dim ItemName As String
    Dim i As Integer
    Dim HasCommands As Boolean
    Dim gID As Integer

    If InMapMode Then
        If tvMapView.SelectedItem Is Nothing Then Exit Sub
    Else
        If tvMenus.SelectedItem Is Nothing Then Exit Sub
    End If
    
    If Not Force Then
        If Not tvMenus.SelectedItem Is Nothing Then
            If IsTemplate(tvMenus.SelectedItem.key) And Not RemoveTemplate Then Exit Sub
        End If
    End If
    
    If InMapMode Then
        If IsTBMapSel Then
            If ToolbarIndexByKey(tvMapView.SelectedItem.key) = 0 Then
                LastSelNode = ""
            Else
                Dim tb As ToolbarDef
                tb = Project.Toolbars(ToolbarIndexByKey(tvMapView.SelectedItem.key))
                
                Ans = MsgBox("Would you like to delete all the items under the '" + tvMapView.SelectedItem.Text + "' toolbar?", vbQuestion + vbYesNo, GetLocalizedStr(546))
                If Ans = vbYes Then
                    Screen.MousePointer = vbHourglass
                    For i = 1 To UBound(tb.Groups)
                        FloodPanel.caption = "Deleting: " + tb.Groups(i)
                        FloodPanel.Value = (i / UBound(tb.Groups)) * 100
                        
                        SelectItem tvMenus.Nodes("G" & GetIDByName(tb.Groups(i))), True, False
                        RemoveItem True, , True
                        DoEvents
                    Next i
                    Screen.MousePointer = vbDefault
                End If
                
                ItemName = GetLocalizedStr(323) + " " + tvMapView.SelectedItem.Text
                For i = ToolbarIndexByKey(tvMapView.SelectedItem.key) To UBound(Project.Toolbars) - 1
                    Project.Toolbars(i) = Project.Toolbars(i + 1)
                Next i
                ReDim Preserve Project.Toolbars(UBound(Project.Toolbars) - 1)
                If UBound(Project.Toolbars) = 0 Then Project.RemoveImageAutoPosCode = False
                RefreshMap
                SaveState GetLocalizedStr(201) + " " + ItemName
            End If
            GoTo CleanUp
        End If
    End If
    
    If Not tvMenus.SelectedItem Is Nothing Then
        If IsGroup(tvMenus.SelectedItem.key) Then
            If Not Force Then
                gID = GetID
                For i = 1 To UBound(MenuCmds)
                    If MenuCmds(i).parent = gID Then
                        HasCommands = True
                        Exit For
                    End If
                Next i
                If HasCommands Then
                    Ans = MsgBox(GetLocalizedStr(273), vbQuestion + vbYesNo, GetLocalizedStr(546))
                Else
                    Ans = vbYes
                End If
            Else
                Ans = vbYes
            End If
            Select Case Ans
                Case vbNo
                    Exit Sub
                Case vbYes
                    ItemName = GetLocalizedStr(270) + " " + tvMenus.SelectedItem.Text
                    RemoveGroupElement GetID, IsSHIFT Or ForceRecurse, True
    '                If CanBeRemoved(GetID) Then
    '                    ItemName = GetLocalizedStr(270) + " " + tvMenus.SelectedItem.Text
    '                    RemoveGroupElement GetID
    '                Else
    '                    MsgBox GetLocalizedStr(274), vbInformation + vbOKOnly, GetLocalizedStr(547)
    '                    Exit Sub
    '                End If
            End Select
        Else
            ItemName = GetLocalizedStr(271) + " " + tvMenus.SelectedItem.Text
            RemoveCommandElement GetID, IsSHIFT Or ForceRecurse, True
        End If
    
        tvMenus.Nodes.Remove tvMenus.SelectedItem.Index
    End If
    
CleanUp:
    
    If Not ForceRecurse Then
        UpdateControls
        If InMapMode Then
            LastSelNode = ""
            RefreshMap
            SetCtrlFocus tvMapView
        Else
            SetCtrlFocus tvMenus
        End If
        
        SaveState GetLocalizedStr(201) + " " + ItemName
    End If

End Sub

Private Sub SetCtrlFocus(ctrl As Control)

    On Error Resume Next
    ctrl.SetFocus

End Sub

'Private Function CanBeRemoved(g As Integer) As Boolean
'
'    Dim c As Integer
'
'    For c = 1 To UBound(MenuCmds)
'        If MenuCmds(c).Actions.onclick.Type = atcCascade And MenuCmds(c).Actions.onclick.TargetMenu = g Or _
'            MenuCmds(c).Actions.onmouseover.Type = atcCascade And MenuCmds(c).Actions.onmouseover.TargetMenu = g Or _
'            MenuCmds(c).Actions.OnDoubleClick.Type = atcCascade And MenuCmds(c).Actions.OnDoubleClick.TargetMenu = g Then
'            CanBeRemoved = False
'            Exit Function
'        End If
'    Next c
'
'    CanBeRemoved = True
'
'End Function

Private Sub ShowMarginsDlg()

    frmGrpMargins.Show vbModal
    
    UpdateLivePreview

End Sub

Private Sub DoUnlock()

    Dim c As String
    Dim sm As Integer
    Dim cm As Integer
    Dim m As Integer
    
    On Error GoTo ReportError

    If Project.HasChanged Then
        If Not NewMenu Then Exit Sub
    End If

    If Val(GetSetting(App.EXEName, "RegInfo", "CacheSig02", "0")) = CacheSignature Then
        sm = CInt(GetSetting(App.EXEName, "RegInfo", "CacheSig01", "0"))
        If sm > 0 Then
            cm = Month(Now)
            m = cm - sm
            If m < 0 Then m = m + 12
            If m < 3 Then
                c = GetSetting(App.EXEName, "RegInfo", "CacheData", "")
                If LenB(c) <> 0 Then
                    c = Inflate(HEX2Str(c))
                    If Split(HEX2Str(c), "|")(4) = GetAppType Then
                        If MsgBox(GetLocalizedStr(694), vbYesNo + vbInformation, GetLocalizedStr(695)) = vbYes Then
                            If RunShellExecute("Open", "register.exe", c, Long2Short(AppPath), 1) > 32 Then
                                tmrClose.Enabled = True
                                Exit Sub
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    DoSilentValidation = False
    frmNewReg.Show vbModal

    If USER = "DEMO" Then Exit Sub

    'USER = "Xavier Flix S."
    'COMPANY = "xFX JumpStart::Software Division"
    'USERSN = "FC5B8617035C7C6C6C37FA0545BA5733016A"

    DoUnlock2

    Exit Sub
    
ReportError:
    MsgBox "Error " & Err.number & " at line " & Erl & ": " & Err.Description, vbCritical + vbOKOnly

End Sub

Private Sub DoUnlock2()

    Dim p As String
    Dim c As String

    p = USER + "|" + COMPANY + "|" + USERSN + "|" + GetHDSerial + "|"
    #If LITE = 0 Then
        #If DEVVER = 0 Then
            p = p + "STD"
        #Else
            p = p + "DEV"
        #End If
    #Else
        p = p + "LIT"
    #End If
    p = p + "|"
    c = Str2HEX(p)
    If RunShellExecute("Open", "register.exe", c, Long2Short(AppPath), 1) <= 32 Then
        MsgBox GetLocalizedStr(696), vbCritical + vbOKOnly, GetLocalizedStr(697)
    Else
        c = Str2HEX(Deflate(c))
        SaveSetting App.EXEName, "RegInfo", "CacheData", c
        SaveSetting App.EXEName, "RegInfo", "CacheSig01", Month(Now)
        SaveSetting App.EXEName, "RegInfo", "CacheSig02", CacheSignature
        
        If IsWinNT Then
            c = c + "|" + CStr(Month(Now)) + "|" + CStr(CacheSignature)
            Compress AppPath + "rsc/reg.dat", c, False
        End If
        
        tmrClose.Enabled = True
    End If

End Sub

Private Sub ChangeLivePreviewState()

    mnuToolsLivePreview.Checked = Not mnuToolsLivePreview.Checked
    Preferences.UseLivePreview = mnuToolsLivePreview.Checked
    
    If Preferences.UseLivePreview Then UpdateLivePreview
   
    Form_Resize

End Sub

'Private Sub mnuToolsPublish_Click()
'
'    PublishMenus
'
'End Sub

'Private Sub PublishMenus()
'
'    Dim OriginalProject As ProjectDef
'
'    If Project.UserConfigs(Project.DefaultConfig).Type = ctcLocal Then
'        MsgBox GetLocalizedStr(678), vbInformation + vbOKOnly, StrConv(GetLocalizedStr(590), vbProperCase)
'        Exit Sub
'    End If
'
'    OriginalProject = Project
'
'    If SomethingIsWrong(False, False) Or _
'        InvalidItemsNames Or _
'        DuplicatedItemsNames Or _
'        InvalidItemsData Or _
'        MissingParameters Then Exit Sub
'
'    If Project.FTP.RemoteInfo4FTP <> "" Then
'        Project.DefaultConfig = GetConfigID(Project.FTP.RemoteInfo4FTP)
'    End If
'
'    If DoCompile Then
'        frmFTPPublishing.Show vbModal
'        If ShowReport Then mnuToolsReport_Click
'    End If
'
'AbortCompile:
'    Project.CodeOptimization = OriginalProject.CodeOptimization
'    Project.DefaultConfig = OriginalProject.DefaultConfig
'
'End Sub

Friend Function DoCompile(Optional IsForPreviewing As Boolean, Optional DoNotCompile As Boolean, Optional vPreviewPath As String, Optional idxTB As Integer, Optional idxGrp As Integer, Optional idxcmd As Integer, Optional IsLivePreview As Boolean) As Boolean

    FixLinks
    If Not IsForPreviewing Then
        If (Preferences.VerifyLinksOptions.VerifyOptions And lvcVerifyWhenCompiling) = lvcVerifyWhenCompiling Then
            LinkVerifyMode = spmcAuto
            BrokenLinksReport
        End If
    End If
    
    If Not IsLivePreview Then
        Me.Enabled = False
        Screen.MousePointer = vbHourglass
    End If
    
    LoadAddInParams Project.AddIn.Name
    
    TipsSys.CanDisable = IsForPreviewing
    tmrDEMOInfo.Enabled = IsDEMO
    If UBound(Project.Toolbars) = 0 Then Project.RemoveImageAutoPosCode = False
    #If LITE = 1 Then
        ApplyLITELimitations
        If UBound(Project.Toolbars) > 1 Then
            ShowLITELImitationInfo 1
            GoTo ExitSub
        End If
    #End If

    sbDummy.Style = sbrSimple
'    If (CacheSignature = Val(GetSetting(App.EXEName, "RegInfo", "CacheSig02", "0")) And Not IsDEMO) Or IsDEMO Then
        DoCompile = CompileProject(MenuGrps, MenuCmds, Project, Preferences, params, IsForPreviewing, DoNotCompile, vPreviewPath, , idxTB, idxcmd, idxGrp, IsLivePreview)
'    Else
'        tmrDoReg.Enabled = False
'        tmrDoReg.Enabled = True
'    End If
    sbDummy.Style = sbrNormal
    
ExitSub:
    With Project
        .DOMCode = ""
        .DOMFramesCode = ""
        .NSCode = ""
        .NSFramesCode = ""
    End With
    
    If Not IsLivePreview Then
        Screen.MousePointer = vbDefault
        Me.Enabled = True
    End If

End Function

#If LITE = 0 Then
Private Sub ShowAddInEditor()

    ShowAIEWarning
    frmAddInEditor.Show

End Sub

Friend Sub ShowAIEWarning()

    With TipsSys
        Do While .IsVisible
            DoEvents
        Loop
        .TipTitle = "AddIn Editor"
        .Tip = GetLocalizedStr(415)
        .CanDisable = True
        .Show
    End With

End Sub
#End If

Private Sub ShowTBEditor()

    Dim selNode As Node
    
    If InMapMode Then
        Set selNode = tvMapView.SelectedItem
    End If

    frmTBEditor.Show vbModal
    If Project.HasChanged Then
        UpdateTitleBar
        UpdateTargetFrameCombo
        UpdateControls
        If InMapMode Then RefreshMap
    Else
        DoLivePreview
    End If
    
End Sub

Private Sub mnuToolsExtractIcon_Click()

    #If LITE = 1 Then
        ShowLITELImitationInfo 2
    #Else
        frmXIcon.Show vbModal
    #End If

End Sub

Private Sub mnuToolsHotSpotsEditor_Click()

    ShowHotSpotsEditor

End Sub

Private Sub mnuToolsIHW_Click()

    #If LITE = 1 Then
        ShowLITELImitationInfo 2
    #Else
        Dim IsInMapViewMode As Boolean
        IsInMapViewMode = Preferences.UseMapView
        If Not IsInMapViewMode Then lblTSViewsMap_Click
        
        frmItemHighlightWizard.Show vbModal
        
        If Not IsInMapViewMode Then lblTSViewsNormal_Click
    #End If

End Sub

Private Sub mnuToolsInstallMenus_Click()

    ReDim SelSecProjects(0)
    SecProjMode = spmcFromInstallMenus
    frmLCMan.Show vbModal

End Sub

Private Sub mnuToolsInstallMenusAIFLC_Click()

    frmLCFramesInstall.Show vbModal

End Sub

Private Sub mnuToolsInstallMenusAILC_Click()

    frmLCInstall.Show vbModal

End Sub

Private Sub mnuToolsInstallMenusAIRLC_Click()

    frmLCRemove.Show vbModal

End Sub

Private Sub mnuToolsLivePreview_Click()

    ChangeLivePreviewState

End Sub

Private Sub mnuToolsPreview_Click()

    ShowPreview

End Sub

Private Sub mnuToolsPublish_Click()

    ToolsPublish

End Sub

Private Sub mnuToolsReport_Click()

    frmCompilationReport.Show vbModal

End Sub

Private Sub mnuToolsSecProj_Click()

    #If LITE = 1 Then
        ShowLITELImitationInfo 2
    #Else
        SecProjMode = spmcFromStdDlg
        frmSecProjDef.Show vbModal
    #End If

End Sub

Private Sub mnuToolsSetDefaultBrowser_Click()

    ShowBrowsersDlg

End Sub

Private Sub mnuToolsToolbarsEditor_Click()

    TBEPage = tbepcGeneral
    ShowTBEditor

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)
    If IsResizing Or IsLoadingProject Then Exit Sub

    Cancel = HandleSubclassMsg(hwnd, uMsg, wParam, lParam)
End Sub

Private Sub opAlignmentStyle_Click(Index As Integer)

    If IsUpdating Then Exit Sub
    
    DontRefreshMap = True
    UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(668)
    
    AdjustMenusAlignment
    
    If InMapMode Then
        SetCtrlFocus tvMapView
    Else
        SetCtrlFocus tvMenus
    End If
    
    UpdateLivePreview

End Sub

Private Sub AdjustMenusAlignment()

    Dim i As Integer
    Dim g As MenuGrp
    Dim gID As Integer
    
    gID = GetID
    g = MenuGrps(GetID)
    
    For i = 1 To UBound(MenuCmds)
        If MenuCmds(i).parent = gID Then
            With MenuCmds(i)
                If .Actions.onmouseover.Type = atcCascade Then
                    If g.AlignmentStyle = ascHorizontal Then
                        Select Case .Actions.onmouseover.TargetMenuAlignment
                            Case gacLeftBottom, gacLeftCenter, gacLeftTop, gacRightBottom, gacRightCenter, gacRightTop
                                .Actions.onmouseover.TargetMenuAlignment = gacBottomLeft
                        End Select
                    Else
                        Select Case .Actions.onmouseover.TargetMenuAlignment
                            Case gacBottomLeft, gacBottomCenter, gacBottomRight, gacTopCenter, gacTopLeft, gacTopRight
                                .Actions.onmouseover.TargetMenuAlignment = gacRightTop
                        End Select
                    End If
                End If
            End With
        End If
    Next i

End Sub

Private Sub picLeading_GotFocus()

    If InMapMode Then
        SetCtrlFocus tvMapView
    Else
        SetCtrlFocus tvMenus
    End If

End Sub

Private Sub picLeading_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MimicHoverButtonHover True

End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplitting = True

End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim newVal As Long
    Static LastVal As Long

    If IsSplitting Then
    
        newVal = picSplit.Left + x
        
        If newVal = LastVal Then Exit Sub
        If newVal < 2633 Then Exit Sub
        
        If Preferences.UseEasyActions Then
            If (Width - newVal) < 3652 + 70 * 15 Then Exit Sub
        Else
            If (Width - newVal) < 3652 Then Exit Sub
        End If
        
        LastVal = newVal
    
        picSplit.Left = newVal
        Form_Resize
    End If

End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplitting = False

End Sub

Private Sub picSplit2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplitting = True

End Sub

Private Sub picSplit2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim newVal As Long
    Static LastVal As Long

    If IsSplitting Then
    
        newVal = picSplit2.Top + y
        
        If newVal = LastVal Then Exit Sub
        If newVal < 2945 Then Exit Sub
        
        If (Height - newVal) < 1875 Then Exit Sub
        
        LastVal = newVal
    
        picSplit2.Top = newVal
        Form_Resize
    End If

End Sub

Private Sub picSplit2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsSplitting = False

End Sub

Private Sub sbDummy_PanelDblClick(ByVal Panel As MSComctlLib.Panel)

    If Panel.key = "sbConfig" Then ReSetDefaultConfig

End Sub

Private Sub NewEmptyProject()

    If NewMenu Then
        If Preferences.ShowPPOnNewProject Then
            ShowProjectProperties
        End If
    End If

End Sub

Private Sub FinishRenaming()

    If IsRenaming Then
        If InMapMode Then SendKeys "{ENTER}", True
        AutoNameItem
    End If

End Sub

Private Sub ShowBrowsersDlg()

    frmBrowsers.Show vbModal
    SetDefBrowserIcon

End Sub

Private Sub ReSetDefaultConfig()

    frmConfigSelect.Show vbModal
    UpdateItemsLinks
    UpdateControls

End Sub

Private Sub DoFind()

    With frmFind
        .Show vbModeless, Me
        .SwitchToFindMode
    End With

End Sub

Private Sub OpenRecent(key As String)

    Dim i As Integer
    Dim tmpRecent As RecentFile
    Dim Index As Integer

    Index = Val(Right(key, 1))
    If Index > 0 Then
        If LoadMenu(RecentFiles(Index).Path, True) Then
            tmpRecent = RecentFiles(Index)
            For i = Index To 2 Step -1
                RecentFiles(i) = RecentFiles(i - 1)
            Next i
            RecentFiles(1) = tmpRecent
        Else
            For i = Index To UBound(RecentFiles) - 1
                RecentFiles(i) = RecentFiles(i + 1)
            Next i
            RecentFiles(UBound(RecentFiles)).Title = ""
            RecentFiles(UBound(RecentFiles)).Path = ""
        End If
        UpdateRecentFilesMenu
    End If
    
End Sub

Private Sub sbEasyBookmark_Click()

    cmdBookmark_Click
    UpdateEasyControls

End Sub

Private Sub sbEasyLink_Click()

    Dim IsG As Boolean
    Dim IsC As Boolean
    Dim IsS As Boolean
    
    If tvMenus.SelectedItem Is Nothing Or IsTBMapSel Then
        Exit Sub
    Else
        IsG = IsGroup(tvMenus.SelectedItem.key)
        IsC = IsCommand(tvMenus.SelectedItem.key)
        IsS = IsSeparator(tvMenus.SelectedItem.key)
    End If
    
    If IsS Then Exit Sub
    
    tsCmdType.Tabs(2).Selected = True
    tsCmdType_Click
    
    If IsC Then
        With MenuCmds(GetID).Actions.onclick
            If .Type = atcNone Then cmbActionType.ListIndex = 1
        End With
    End If
    
    If IsG Then
        With MenuGrps(GetID).Actions.onclick
            If .Type = atcNone Then cmbActionType.ListIndex = 1
        End With
    End If

    cmdBrowse_Click
    UpdateEasyControls

End Sub

Private Sub sbEasyNewWindow_Click()

    If chkEasyEnableNewWindow.Value = vbUnchecked Then
        chkEasyEnableNewWindow.Value = vbChecked
    End If

    cmdWinParams_Click
    UpdateEasyControls

End Sub

Private Sub sbEasyTargetFrame_Click()

    cmdTargetFrame_Click
    UpdateEasyControls

End Sub

Private Sub sbF1_Click()

    'Dim a(1) As Variant

    'a(0) = vbKeyF1
    'a(1) = 0
    'CallByName Screen.ActiveForm, "Form_KeyDown", VbMethod, a
    DisplayTip "Dialog Help", "Did you know that you can access the 'Dialog Help' for every single dialog in DHTML Menu Builder by pressing the F1 key?", True
    SendKeys "{F1}"

End Sub

Private Sub sbLP_Close_Click()

    ChangeLivePreviewState

End Sub

Private Sub sbLP_GroupMode_Click()

    If sbLP_GroupMode.tag = "TBI" Then
        sbLP_GroupMode.tag = "GRP"
        sbLP_GroupMode.BackColor = &H8000000E
    Else
        sbLP_GroupMode.tag = "TBI"
        sbLP_GroupMode.BackColor = &H8000000F
    End If
    
    DoDelayedLivePreview

End Sub

Private Sub tbCmd_ButtonClick(ByVal Button As MSComctlLib.Button)

    Dim IsG As Boolean
    
    FinishRenaming
    IsRenaming = False
    
    IsG = IsGroup(tvMenus.SelectedItem.key)
    
    tmrInitStyleDlg.Enabled = True

    Select Case Button.key
        Case "tbColor"
            SelectiveLauncher = slcSelColor
            If IsG Then
                frmGrpColor.Show vbModal
            Else
                frmColor.Show vbModal
            End If
        Case "tbCursor"
            SelectiveLauncher = slcSelCursor
            If IsG Then
                frmGrpCursor.Show vbModal
            Else
                frmCursor.Show vbModal
            End If
        Case "tbFont"
            SelectiveLauncher = slcSelFont
            If IsG Then
                frmGrpFont.Show vbModal
            Else
                frmFont.Show vbModal
            End If
        Case "tbImage"
            SelectiveLauncher = slcSelImage
            If IsG Then
                frmGrpImage.Show vbModal
            Else
                frmImage.Show vbModal
            End If
        Case "tbMargins"
            SelectiveLauncher = slcSelMargin
            frmGrpMargins.Show vbModal
        Case "tbFX"
            SelectiveLauncher = slcSelSFX
            ShowSFXDlg
        Case "tbSelFX"
            SelectiveLauncher = slcSelSelFX
            If IsG Then
                frmGrpSelectionFX.Show vbModal
            Else
                frmSelFX.Show vbModal
            End If
        Case "tbUp"
            tmrInitStyleDlg.Enabled = False
            
            LockWindowUpdate Me.hwnd
            MoveItemUp
            LockWindowUpdate 0
            
            Exit Sub
        Case "tbDown"
            tmrInitStyleDlg.Enabled = False
            
            LockWindowUpdate Me.hwnd
            MoveItemDown
            LockWindowUpdate 0
            
            Exit Sub
        Case "tbSound"
           'frmSound.Show vbModal
    End Select
    
    UpdateLivePreview
    
End Sub

Private Sub MoveItemUp()

    Dim t As Integer
    Dim gID As Integer
    Dim tmpName As String
    Dim g As Integer
    Dim tmpTB As ToolbarDef
    
    If IsUpdating Or IsRefreshingMap Or IsRestoringExp Then Exit Sub
    
    SaveCurExp

    If InMapMode And IsTBMapSel Then
        t = ToolbarIndexByKey(tvMapView.SelectedItem.key)
        If t <= 1 Then Exit Sub
        With Project
            tmpTB = .Toolbars(t)
            .Toolbars(t) = .Toolbars(t - 1)
            .Toolbars(t - 1) = tmpTB
        End With
    Else
        If IsGroup(tvMenus.SelectedItem.key) Then
            MoveGroupUp
            gID = GetID
            For t = 1 To UBound(Project.Toolbars)
                If MemberOf(gID) = t Then
                    tmpName = MenuGrps(gID).Name
                    With Project.Toolbars(t)
                        For g = 2 To UBound(.Groups)
                            If .Groups(g) = tmpName Then
                                .Groups(g) = .Groups(g - 1)
                                .Groups(g - 1) = tmpName
                                Exit For
                            End If
                        Next g
                    End With
                End If
            Next t
            SaveState GetLocalizedStr(275) + " " + MenuGrps(GetID).Name + " " + GetLocalizedStr(276)
        Else
            If Not MoveUp Then Exit Sub
            SaveState GetLocalizedStr(275) + " " + MenuCmds(GetID).Name + " " + GetLocalizedStr(276)
        End If
    End If
    
    If InMapMode Then
        RefreshMap False
        
        RestoreCurExp
        
        tvMapView.SelectedItem.EnsureVisible
    End If

End Sub

Private Sub MoveItemDown()

    Dim t As Integer
    Dim gID As Integer
    Dim tmpName As String
    Dim g As Integer
    Dim tmpTB As ToolbarDef
    
    If IsUpdating Or IsRefreshingMap Or IsRestoringExp Then Exit Sub

    SaveCurExp

    If InMapMode And IsTBMapSel Then
        t = ToolbarIndexByKey(tvMapView.SelectedItem.key)
        With Project
            If t = UBound(.Toolbars) Then Exit Sub
            tmpTB = .Toolbars(t)
            .Toolbars(t) = .Toolbars(t + 1)
            .Toolbars(t + 1) = tmpTB
        End With
    Else
        If IsGroup(tvMenus.SelectedItem.key) Then
            MoveGroupDown
            gID = GetID
            For t = 1 To UBound(Project.Toolbars)
                If MemberOf(gID) = t Then
                    tmpName = MenuGrps(gID).Name
                    With Project.Toolbars(t)
                        For g = 1 To UBound(.Groups) - 1
                            If .Groups(g) = tmpName Then
                                .Groups(g) = .Groups(g + 1)
                                .Groups(g + 1) = tmpName
                                Exit For
                            End If
                        Next g
                    End With
                End If
            Next t
            SaveState GetLocalizedStr(275) + " " + MenuGrps(GetID).Name + " " + GetLocalizedStr(277)
        Else
            If Not MoveDown Then Exit Sub
            SaveState GetLocalizedStr(275) + " " + MenuCmds(GetID).Name + " " + GetLocalizedStr(277)
        End If
    End If
    
    If InMapMode Then
        RefreshMap False
        
        RestoreCurExp
        
        tvMapView.SelectedItem.EnsureVisible
    End If

End Sub

Private Sub SaveCurExp()

    Dim i As Integer
    Dim n As Node
    
    If Not InMapMode Then Exit Sub
    
    On Error GoTo ExitSub

    ReDim nodexExp(tvMapView.Nodes.Count)
    i = 0
    For Each n In tvMapView.Nodes
        If n.Expanded Then
            nodexExp(i) = n.FullPath
            i = i + 1
        End If
    Next n
    
    ReDim Preserve nodexExp(i)
    
ExitSub:

End Sub

Private Sub RestoreCurExp()

    Dim i As Integer
    Dim n As Node
    Dim State As Boolean
    
    If Not InMapMode Then Exit Sub
    
    On Error GoTo ExitSub
    
    IsRestoringExp = True

    For Each n In tvMapView.Nodes
        State = False
        For i = 0 To UBound(nodexExp) - 1
            If nodexExp(i) = n.FullPath Then
                State = True
                Exit For
            End If
        Next i
        If n.Expanded <> State Then n.Expanded = State
    Next n

    RefreshMap

ExitSub:

    IsRestoringExp = False

End Sub

Private Sub tbCmd_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        Set SelToolbar = tbCmd
        'PopupMenu mnuToolbars, vbRightButton, X, Y + SelToolbar.Top
    End If

End Sub

Private Sub tbCmd_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MimicHoverButtonHover False

End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)

    FinishRenaming
    IsRenaming = False

    Select Case Button.key
        Case "tbNew"
            NewEmptyProject
        Case "tbOpen"
            LoadMenu
        Case "tbSave"
            SaveMenu
        Case "tbAddGrp"
            MenuAddGroup
        Case "tbAddCmd"
            MenuAddCommand
        Case "tbAddSep"
            MenuAddSeparator
        Case "tbRemove"
            RemoveItem
        Case "tbCopy"
            SelectiveCopy
        Case "tbPaste"
            SelectivePaste
        'Case "tbHelp"
        '    Me.WhatsThisMode
        Case "tbUndo"
            DoUndo
        Case "tbRedo"
            DoRedo
        Case "tbFind"
            DoFind
        Case "tbAddSubGrp"
            MenuAddSubGroup
    End Select

End Sub

Private Sub tbMenu_ButtonDropDown(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "tbOpen"
            PopupMenu mnuFileOpenRecent, , tbMenu.Left + Button.Left, tbMenu.Top + tbMenu.Height
        Case "tbUndo"
            SelRedoUndo = sUndo
            ShowUndoRedoList
        Case "tbRedo"
            SelRedoUndo = sRedo
            ShowUndoRedoList
        Case "tbNew"
            PopupMenu mnuFileNew, , tbMenu.Left + Button.Left, tbMenu.Top + tbMenu.Height
    End Select

End Sub

Private Sub ShowUndoRedoList()

    Dim i As Integer
    
    With frmUndoRedoStates
        .lstStates.Clear
        .txtMode.Text = ""
        Select Case SelRedoUndo
            Case sUndo
                For i = CurState To 1 Step -1
                    .lstStates.AddItem UndoStates(i).Description
                Next i
                .Move Left + tbMenu.Left + tbMenu.Buttons("tbUndo").Left + 70, Top + tbMenu.Top + tbMenu.Height + 635
            Case sRedo
                For i = UBound(UndoStates) To CurState + 1 Step -1
                    .lstStates.AddItem UndoStates(i).Description
                Next i
                .Move Left + tbMenu.Left + tbMenu.Buttons("tbRedo").Left + 70, Top + tbMenu.Top + tbMenu.Height + 635
        End Select
        .Show vbModal
    End With
    
    Select Case SelRedoUndo
        Case sUndo
            CurState = CurState - SelRedoUndoCount
            DoUndo
        Case sRedo
            CurState = CurState + SelRedoUndoCount
            DoRedo
    End Select

End Sub

Private Sub tbMenu_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        Set SelToolbar = tbMenu
        'PopupMenu mnuToolbars, vbRightButton, X, Y + SelToolbar.Top
    End If

End Sub

Private Sub tbMenu_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MimicHoverButtonHover False

End Sub

Private Sub tbMenu2_ButtonClick(ByVal Button As MSComctlLib.Button)

    FinishRenaming
    IsRenaming = False

    Select Case Button.key
        Case "tbCompile"
            ToolsCompile
        Case "tbPublish"
            ToolsPublish
        Case "tbPreview"
            ShowPreview
        Case "tbProperties"
            ProjectPropertiesPage = pppcGeneral
            ShowProjectProperties
        Case "tbHotSpotEditor"
            ShowHotSpotsEditor
        Case "tbTBEditor"
            TBEPage = tbepcGeneral
            ShowTBEditor
    End Select

End Sub

Private Sub ShowHotSpotsEditor()
    
    Dim ThisConfig As ConfigDef
    
    On Error GoTo ExitWithError
    
    If Project.UserConfigs(Project.DefaultConfig).Type = ctcLocal Then
        DisplayTip GetLocalizedStr(686), GetLocalizedStr(687)
    End If
    
    If CreateToolbar Then
        If MsgBox(GetLocalizedStr(278) + vbCrLf + GetLocalizedStr(279) + vbCrLf + GetLocalizedStr(280) + vbCrLf + vbCrLf + GetLocalizedStr(548), vbInformation + vbYesNo, GetLocalizedStr(335)) = vbNo Then Exit Sub
    End If
    
    If SomethingIsWrong(False, True) Then Exit Sub
    
    ThisConfig = Project.UserConfigs(Project.DefaultConfig)

    'Verify that the destination file is not read-only
    If GetAttr(ThisConfig.HotSpotEditor.HotSpotsFile) And vbReadOnly Then
        MsgBox GetLocalizedStr(281), vbInformation + vbOKOnly, GetLocalizedStr(549)
        Exit Sub
    End If
    
    Project.RemoveImageAutoPosCode = False
    
    frmHotSpotsEditor2.Show vbModal
    'If Not HSCanceled And ThisConfig.Frames.UseFrames Then
    '    frmLCFramesInstall.Show vbModal
    'End If
   
    Exit Sub
    
ExitWithError:
    MsgBox GetLocalizedStr(557) + vbCrLf + GetLocalizedStr(544) + " " & Err.number & ": " + Err.Description, vbOKOnly, GetLocalizedStr(551)

End Sub

Private Function InvalidItemsNames() As Boolean

    Dim j As Integer
    Dim i As Integer
    Dim m As String
    Dim InvalidChars As String
    Dim CommMsg As String
    
    InvalidChars = "~!@#$%^&*()+={}[]:;""'<>,./?\|"
    CommMsg = GetLocalizedStr(558) + " "
    
    For j = 1 To UBound(MenuGrps)
        If LenB(MenuGrps(j).Name) = 0 Then
            MenuGrps(j).Name = GetSecuenceName(True, "Group")
            tvMenus.Nodes("G" & j).Text = MenuGrps(j).Name
            RefreshMap
        End If
        For i = 1 To Len(InvalidChars)
            m = Mid$(InvalidChars, i, 1)
            If InStr(MenuGrps(j).Name, m) Then
                InvalidItemsNames = True
                MsgBox CommMsg + GetLocalizedStr(559) + ": '" + MenuGrps(j).Name + "' " + GetLocalizedStr(561), vbCritical + vbOKOnly, "Error Compiling"
                Exit Function
            End If
        Next i
    Next j
    
    For j = 1 To UBound(MenuCmds)
        If LenB(MenuCmds(j).Name) = 0 Then
            MenuCmds(j).Name = GetSecuenceName(True, "Command")
            tvMenus.Nodes("C" & j).Text = MenuCmds(j).Name
            RefreshMap
        End If
        If MenuCmds(j).Name <> "[SEP]" Then
            For i = 1 To Len(InvalidChars)
                m = Mid$(InvalidChars, i, 1)
                If InStr(MenuCmds(j).Name, m) Then
                    InvalidItemsNames = True
                    MsgBox CommMsg + GetLocalizedStr(560) + ": '" + MenuCmds(j).Name + "' " + GetLocalizedStr(561), vbCritical + vbOKOnly, "Error Compiling"
                    Exit Function
                End If
            Next i
        End If
    Next j

End Function

Private Function InvalidItemsData() As Boolean

    Dim j As Integer
    Dim i As Integer
    Dim m As String
    Dim InvalidChars As String
    Dim CommMsg As String
    
    InvalidChars = """"
    CommMsg = GetLocalizedStr(558) + " "
    
    For j = 1 To UBound(MenuCmds)
        If MenuCmds(j).Name <> "[SEP]" Then
            For i = 1 To Len(InvalidChars)
                m = Mid$(InvalidChars, i, 1)
                If InStr(MenuCmds(j).Actions.onclick.url, m) And Not UsesProtocol(MenuCmds(j).Actions.onclick.url) Then
                    InvalidItemsData = True
                    MsgBox CommMsg + GetLocalizedStr(562) + ": '" + MenuCmds(j).Name + "' " + GetLocalizedStr(561), vbCritical + vbOKOnly, GetLocalizedStr(569)
                    Exit Function
                End If
                If InStr(MenuCmds(j).Actions.onmouseover.url, m) And Not UsesProtocol(MenuCmds(j).Actions.onmouseover.url) Then
                    InvalidItemsData = True
                    MsgBox CommMsg + GetLocalizedStr(563) + ": '" + MenuCmds(j).Name + "' " + GetLocalizedStr(561), vbCritical + vbOKOnly, GetLocalizedStr(569)
                    Exit Function
                End If
                If InStr(MenuCmds(j).Actions.OnDoubleClick.url, m) And Not UsesProtocol(MenuCmds(j).Actions.OnDoubleClick.url) Then
                    InvalidItemsData = True
                    MsgBox CommMsg + GetLocalizedStr(564) + ": '" + MenuCmds(j).Name + "' " + GetLocalizedStr(561), vbCritical + vbOKOnly, GetLocalizedStr(569)
                    Exit Function
                End If
            Next i
        End If
    Next j
    
    For j = 1 To UBound(MenuGrps)
        For i = 1 To Len(InvalidChars)
            m = Mid$(InvalidChars, i, 1)
            If InStr(MenuGrps(j).Actions.onclick.url, m) And Not UsesProtocol(MenuGrps(j).Actions.onclick.url) Then
                InvalidItemsData = True
                MsgBox CommMsg + GetLocalizedStr(565) + ": '" + MenuGrps(j).Name + "' " + GetLocalizedStr(561), vbCritical + vbOKOnly, GetLocalizedStr(569)
                Exit Function
            End If
            If InStr(MenuGrps(j).Actions.onmouseover.url, m) And Not UsesProtocol(MenuGrps(j).Actions.onmouseover.url) Then
                InvalidItemsData = True
                MsgBox CommMsg + GetLocalizedStr(566) + ": '" + MenuGrps(j).Name + "' " + GetLocalizedStr(561), vbCritical + vbOKOnly, GetLocalizedStr(569)
                Exit Function
            End If
            If InStr(MenuGrps(j).Actions.OnDoubleClick.url, m) And Not UsesProtocol(MenuGrps(j).Actions.OnDoubleClick.url) Then
                InvalidItemsData = True
                MsgBox CommMsg + GetLocalizedStr(567) + ": '" + MenuGrps(j).Name + "' " + GetLocalizedStr(561), vbCritical + vbOKOnly, GetLocalizedStr(569)
                Exit Function
            End If
        Next i
    Next j
    
    InvalidChars = """"
    
    For j = 1 To UBound(MenuCmds)
        If MenuCmds(j).Name <> "[SEP]" Then
            For i = 1 To Len(InvalidChars)
                m = Mid$(InvalidChars, i, 1)
                If InStr(MenuCmds(j).WinStatus, m) Then
                    InvalidItemsData = True
                    MsgBox CommMsg + GetLocalizedStr(560) + ": '" + MenuCmds(j).Name + "' " + GetLocalizedStr(568), vbCritical + vbOKOnly, GetLocalizedStr(569)
                    Exit Function
                End If
            Next i
        End If
    Next j
    
    If CreateToolbar Then
        For j = 1 To UBound(MenuGrps)
            If MenuGrps(j).IncludeInToolbar Then
                For i = 1 To Len(InvalidChars)
                    m = Mid$(InvalidChars, i, 1)
                    If InStr(MenuGrps(j).WinStatus, m) Then
                        InvalidItemsData = True
                        MsgBox CommMsg + GetLocalizedStr(559) + ": '" + MenuGrps(j).Name + "' " + GetLocalizedStr(568), vbCritical + vbOKOnly, GetLocalizedStr(569)
                        Exit Function
                    End If
                Next i
            End If
        Next j
    End If
    
End Function

Private Function DuplicatedItemsNames() As Boolean

    'Dim c1 As Integer
    'Dim c2 As Integer
    Dim g1 As Integer
    Dim g2 As Integer
    Dim n As String
    
ReStart:
    For g1 = 1 To UBound(MenuGrps)
        For g2 = g1 + 1 To UBound(MenuGrps)
            If MenuGrps(g1).Name = MenuGrps(g2).Name Then
                If InMapMode Then
                    MenuGrps(g1).Name = GetSecuenceName(True, "Group")
                    GoTo ReStart
                Else
                    n = MenuGrps(g1).Name
                    GoTo ExitAndShowErrMsg
                End If
            End If
        Next g2
'        For c1 = 1 To UBound(MenuCmds)
'            If MenuCmds(c1).Name <> "[SEP]" Then
'                For c2 = c1 + 1 To UBound(MenuCmds)
'                    If MenuCmds(c1).Name = MenuCmds(c2).Name Then
'                        If InMapMode Then
'                            MenuCmds(c1).Name = GetSecuenceName(False, "Command")
'                            GoTo ReStart
'                        Else
'                            n = MenuCmds(c1).Name
'                            GoTo ExitAndShowErrMsg
'                        End If
'                    End If
'                Next c2
'                If MenuGrps(g1).Name = MenuCmds(c1).Name Then
'                    If InMapMode Then
'                        MenuCmds(c1).Name = GetSecuenceName(False, "Command")
'                        GoTo ReStart
'                    Else
'                        n = MenuGrps(g1).Name
'                        GoTo ExitAndShowErrMsg
'                    End If
'                End If
'            End If
'        Next c1
    Next g1
    
    Exit Function
    
ExitAndShowErrMsg:
    MsgBox GetLocalizedStr(570) + ": '" + n + "'", vbCritical + vbOKOnly, GetLocalizedStr(569)
    DuplicatedItemsNames = True

End Function

Private Function SomethingIsWrong(IsForPreview As Boolean, IsForHSE As Boolean) As Boolean

    Dim i As Integer
    Dim CompiledPath As String
    Dim ImagesPath As String
    Dim ThisConfig As ConfigDef
    Dim ConfigMsg As String
    Dim aiCode As String
    Dim aiFileName As String
    
    On Error Resume Next

StartChecks:

    ThisConfig = Project.UserConfigs(Project.DefaultConfig)
    If ThisConfig.Type = ctcRemote Then
        CompiledPath = Project.UserConfigs(GetConfigID(ThisConfig.LocalInfo4RemoteConfig)).CompiledPath
        ImagesPath = Project.UserConfigs(GetConfigID(ThisConfig.LocalInfo4RemoteConfig)).ImagesPath
    Else
        CompiledPath = ThisConfig.CompiledPath
        ImagesPath = ThisConfig.ImagesPath
    End If
    
    ConfigMsg = vbCrLf + vbCrLf + GetLocalizedStr(571) + ": " + ThisConfig.Name

    If Not IsForPreview Then
        'Verify that the destination folder is valid Step 1
        If LenB(CompiledPath) = 0 Then
            If MsgBox(GetLocalizedStr(572) + vbCrLf + GetLocalizedStr(573) + ConfigMsg, vbQuestion + vbYesNo, GetLocalizedStr(574)) = vbYes Then
                frmProjProp.Show vbModal
                GoTo StartChecks
            Else
                SomethingIsWrong = True
                Exit Function
            End If
        End If
        
        'Verify that the destination folder is valid Step 2
        On Error Resume Next
        If Not ((GetAttr(CompiledPath) And vbDirectory) = vbDirectory) Then
            MsgBox GetLocalizedStr(575) + ConfigMsg, vbInformation + vbOKOnly, GetLocalizedStr(576)
            SomethingIsWrong = True
            Exit Function
        End If
        
        'Verify that the destination images folder is valid Step 1
        If LenB(ImagesPath) = 0 Then
            If MsgBox(GetLocalizedStr(577) + vbCrLf + GetLocalizedStr(573) + ConfigMsg, vbQuestion + vbYesNo, GetLocalizedStr(578)) = vbYes Then
                frmProjProp.Show vbModal
                GoTo StartChecks
            Else
                SomethingIsWrong = True
                Exit Function
            End If
        End If
        
        'Verify that the destination images folder is valid Step 2
        On Error Resume Next
        If Not ((GetAttr(ImagesPath) And vbDirectory) = vbDirectory) Then
            MsgBox GetLocalizedStr(579) + ConfigMsg, vbInformation + vbOKOnly, GetLocalizedStr(580)
            SomethingIsWrong = True
            Exit Function
        End If
        
        'HotSpot Editor parameters
        If IsForHSE Then
            'Verify that the destination file is valid
            If Not FileExists(ThisConfig.HotSpotEditor.HotSpotsFile) Then
                If MsgBox(GetLocalizedStr(581) + vbCrLf + GetLocalizedStr(573) + ConfigMsg, vbQuestion + vbYesNo, GetLocalizedStr(582)) = vbYes Then
                    frmProjProp.Show vbModal
                    GoTo StartChecks
                Else
                    SomethingIsWrong = True
                    Exit Function
                End If
            End If
            
            'Verify that the destination file is inside the root web
            If ThisConfig.Type <> ctcRemote And LCase$(ThisConfig.RootWeb) <> LCase$(Left$(ThisConfig.HotSpotEditor.HotSpotsFile, Len(ThisConfig.RootWeb))) Then
                If MsgBox(GetLocalizedStr(583) + vbCrLf + GetLocalizedStr(573) + ConfigMsg, vbQuestion + vbYesNo, GetLocalizedStr(584)) = vbYes Then
                    frmProjProp.Show vbModal
                    GoTo StartChecks
                Else
                    SomethingIsWrong = True
                    Exit Function
                End If
            End If
        End If
        
        'Verify that the destination is a folder inside the root web
        If ThisConfig.Type <> ctcRemote And LCase$(ThisConfig.RootWeb) <> LCase$(Left$(CompiledPath, Len(ThisConfig.RootWeb))) Then
            If MsgBox(GetLocalizedStr(587) + vbCrLf + GetLocalizedStr(588) + ConfigMsg, vbQuestion + vbYesNo, GetLocalizedStr(576)) = vbYes Then
                frmProjProp.Show vbModal
                GoTo StartChecks
            Else
                SomethingIsWrong = True
                Exit Function
            End If
        End If
        
        'Verify that rootwebs are correct and valid
        For i = 0 To UBound(Project.UserConfigs)
            With Project.UserConfigs(i)
                Select Case .Type
                    Case ctcLocal Or ctcCDROM
                        If Left$(LCase$(.RootWeb), 2) <> "\\" And Mid$(LCase$(.RootWeb), 2, 2) <> ":\" Then
                            MsgBox GetLocalizedStr(589) + " " + .Name + " " + GetLocalizedStr(590), vbInformation + vbOKOnly, GetLocalizedStr(591)
                            SomethingIsWrong = True
                            Exit Function
                        End If
                    Case ctcRemote
                        If Left$(LCase$(.RootWeb), 7) <> "http://" And Left$(LCase$(.RootWeb), 8) <> "https://" Then
                            MsgBox GetLocalizedStr(592) + " " + .Name + " " + GetLocalizedStr(590), vbInformation + vbOKOnly, GetLocalizedStr(591)
                            SomethingIsWrong = True
                            Exit Function
                        End If
                End Select
            End With
        Next i
        
        'Verify that FTP parameters are correct
        'If Project.FTP.FTPAddress <> "" Then
        '    If Left$(LCase$(Project.FTP.FTPAddress), 6) <> "ftp://" Then
        '        MsgBox GetLocalizedStr(593), vbInformation + vbOKOnly, GetLocalizedStr(591)
        '        SomethingIsWrong = True
        '        Exit Function
        '    End If
        'End If
SkipThisTest:
    End If
    
    'Verify that cascade commands point to valid groups
    For i = 1 To UBound(MenuCmds)
        If MenuCmds(i).iCursor.cType = 0 Then MenuCmds(i).iCursor.cType = iccDefault
        With MenuCmds(i).Actions
            If .onmouseover.Type = atcCascade And (.onmouseover.TargetMenu > UBound(MenuGrps) Or .onmouseover.TargetMenu = 0) Then
                MsgBox StrConv(GetLocalizedStr(560), vbProperCase) + " '" + NiceCmdCaption(i) + "' " + GetLocalizedStr(594) + " '" + GetLocalizedStr(105) + "' " + GetLocalizedStr(595), vbInformation + vbOKOnly, GetLocalizedStr(596)
                SomethingIsWrong = True
                Exit Function
            End If
            If .onclick.Type = atcCascade And (.onclick.TargetMenu > UBound(MenuGrps) Or .onclick.TargetMenu = 0) Then
                MsgBox StrConv(GetLocalizedStr(560), vbProperCase) + " '" + NiceCmdCaption(i) + "' " + GetLocalizedStr(594) + " '" + GetLocalizedStr(106) + "' " + GetLocalizedStr(595), vbInformation + vbOKOnly, GetLocalizedStr(596)
                SomethingIsWrong = True
                Exit Function
            End If
            If .OnDoubleClick.Type = atcCascade And (.OnDoubleClick.TargetMenu > UBound(MenuGrps) Or .OnDoubleClick.TargetMenu = 0) Then
                MsgBox StrConv(GetLocalizedStr(560), vbProperCase) + " '" + NiceCmdCaption(i) + "' " + GetLocalizedStr(594) + " '" + GetLocalizedStr(107) + "' " + GetLocalizedStr(595), vbInformation + vbOKOnly, GetLocalizedStr(596)
                SomethingIsWrong = True
                Exit Function
            End If
        End With
    Next i
    
    For i = 1 To UBound(MenuGrps)
        If MenuGrps(i).iCursor.cType = 0 Then MenuGrps(i).iCursor.cType = iccDefault
        With MenuGrps(i).Actions
            If .onmouseover.Type = atcCascade Then
                If .onmouseover.TargetMenu > UBound(MenuGrps) Then
                    MsgBox StrConv(GetLocalizedStr(559), vbProperCase) + " '" + NiceGrpCaption(i) + "' " + GetLocalizedStr(594) + " '" + GetLocalizedStr(105) + "' " + GetLocalizedStr(595), vbInformation + vbOKOnly, GetLocalizedStr(596)
                    SomethingIsWrong = True
                    Exit Function
                ElseIf .onmouseover.TargetMenu = 0 Then
                    .onmouseover.TargetMenu = i
                End If
            End If
            If .onclick.Type = atcCascade Then
                If .onclick.TargetMenu > UBound(MenuGrps) Then
                    MsgBox StrConv(GetLocalizedStr(559), vbProperCase) + " '" + NiceGrpCaption(i) + "' " + GetLocalizedStr(594) + " '" + GetLocalizedStr(106) + "' " + GetLocalizedStr(595), vbInformation + vbOKOnly, GetLocalizedStr(596)
                    SomethingIsWrong = True
                    Exit Function
                ElseIf .onclick.TargetMenu = 0 Then
                    .onclick.TargetMenu = i
                End If
            End If
            If .OnDoubleClick.Type = atcCascade Then
                If .OnDoubleClick.TargetMenu > UBound(MenuGrps) Then
                    MsgBox StrConv(GetLocalizedStr(559), vbProperCase) + " '" + NiceGrpCaption(i) + "' " + GetLocalizedStr(594) + " '" + GetLocalizedStr(107) + "' " + GetLocalizedStr(595), vbInformation + vbOKOnly, GetLocalizedStr(596)
                    SomethingIsWrong = True
                    Exit Function
                ElseIf .OnDoubleClick.TargetMenu = 0 Then
                    .OnDoubleClick.TargetMenu = i
                End If
            End If
        End With
    Next i
    
    'Verify that AddIn used by the Project exists
    'and its not from an older version
    With Project.AddIn
        If LenB(.Name) <> 0 Then
            aiFileName = AppPath + "AddIns\" + .Name + ".ext"
            If Not FileExists(aiFileName) Then
                MsgBox GetLocalizedStr(407) + " '" + .Name + "' " + GetLocalizedStr(408) + ". " + GetLocalizedStr(597), vbInformation + vbOKOnly, GetLocalizedStr(598)
                SomethingIsWrong = True
                Exit Function
            Else
                On Error Resume Next
                Err.Clears
                aiCode = LoadFile(aiFileName)
                If (InStr(aiCode, "%%PRINTTBS") = 0) Or _
                    (InStr(aiCode, "%%TOOLBARVARS") = 0) Or _
                    (InStr(aiCode, "%%KBDNAVSUP") = 0) Or _
                    (Err.number <> 0) Then
                    MsgBox GetLocalizedStr(407) + " '" + .Name + "' " + GetLocalizedStr(599), vbInformation + vbOKOnly, GetLocalizedStr(600)
                    SomethingIsWrong = True
                    Exit Function
                End If
            End If
        End If
    End With

End Function

Private Sub tbMenu2_ButtonDropDown(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "tbPreview"
            PopupMenu mnuBrowsers, , tbMenu2.Left + Button.Left, tbMenu2.Top + tbMenu2.Height
            DoEvents
        Case "tbProperties"
            PopupMenu mnuPPShortcuts, , tbMenu2.Left + Button.Left, tbMenu2.Top + tbMenu2.Height
            DoEvents
        Case "tbTBEditor"
            PopupMenu mnuTBEShortcuts, , tbMenu2.Left + Button.Left, tbMenu2.Top + tbMenu2.Height
            DoEvents
    End Select

End Sub

Private Sub tbMenu2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbRightButton Then
        Set SelToolbar = tbMenu2
        'PopupMenu mnuToolbars, vbRightButton, X, Y + SelToolbar.Top
    End If

End Sub

Private Sub tbMenu2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MimicHoverButtonHover False

End Sub

Private Function MoveUp() As Boolean

    Dim tmpNode As Node
    Dim tmpCmd As MenuCmd
    Dim SelId As Integer
    Dim pid As Integer
    
    Set tmpNode = tvMenus.SelectedItem
    If tmpNode.Previous Is Nothing Then
        MoveUp = False
        Exit Function
    End If

    SelId = GetID
    pid = GetID(tmpNode.Previous)
    
    tmpNode.key = ""
    tmpNode.Previous.key = ""
    
    tmpCmd = MenuCmds(SelId)
    MenuCmds(SelId) = MenuCmds(pid)
    MenuCmds(pid) = tmpCmd
    
    With tmpNode
        .Text = IIf(MenuCmds(SelId).Name = "[SEP]", String$(10, "-"), MenuCmds(SelId).Name)
        .key = IIf(MenuCmds(SelId).Name = "[SEP]", "S", "C") & SelId
        .Image = GenCmdIcon(SelId)
    End With
    
    With tmpNode.Previous
        .Text = IIf(MenuCmds(pid).Name = "[SEP]", String$(10, "-"), MenuCmds(pid).Name)
        .key = IIf(MenuCmds(pid).Name = "[SEP]", "S", "C") & pid
        .Image = GenCmdIcon(pid)
    End With
    
    SelectItem tmpNode.Previous, , False
    
    MoveUp = True
    
End Function

Private Sub MoveGroupUp()

    Dim tmpNode As Node
    Dim dummy As Node
    Dim tmpGrp As MenuGrp
    Dim SelId As Integer
    Dim pid As Integer
    
    Set tmpNode = tvMenus.SelectedItem
    If tmpNode.Previous Is Nothing Then Exit Sub

    SelId = GetID
    pid = GetID(tmpNode.Previous)
    
    tmpNode.key = ""
    tmpNode.Previous.key = ""
    
    tmpGrp = MenuGrps(SelId)
    MenuGrps(SelId) = MenuGrps(pid)
    MenuGrps(pid) = tmpGrp
    
    With tmpNode
        .Text = MenuGrps(SelId).Name
        .key = "G" & (SelId)
        .Image = GenGrpIcon(SelId)
        .ForeColor = IIf(MenuGrps(SelId).IsTemplate, vbBlue, vbBlack)
    End With
    
    With tmpNode.Previous
        .Text = MenuGrps(pid).Name
        .key = "G" & (pid)
        .Image = GenGrpIcon(pid)
        .ForeColor = IIf(MenuGrps(pid).IsTemplate, vbBlue, vbBlack)
    End With
    
    Set dummy = tvMenus.Nodes.Add(, , "DUMMY")
    Do Until tmpNode.Child Is Nothing
        Set tmpNode.Child.parent = dummy
    Loop
    
    Do Until tmpNode.Previous.Child Is Nothing
        MenuCmds(GetID(tmpNode.Previous.Child.LastSibling)).parent = GetID(tmpNode)
        Set tmpNode.Previous.Child.LastSibling.parent = tmpNode
    Loop
    
    Do Until dummy.Child Is Nothing
        MenuCmds(GetID(dummy.Child)).parent = GetID(tmpNode.Previous)
        Set dummy.Child.parent = tmpNode.Previous
    Loop
    tvMenus.Nodes.Remove "DUMMY"
    
    FixTargetMenus pid, SelId
    
    SelectItem tmpNode.Previous, , False
    
End Sub

Private Sub MoveGroupDown()

    Dim tmpNode As Node
    Dim dummy As Node
    Dim tmpGrp As MenuGrp
    Dim SelId As Integer
    Dim pid As Integer
    
    Set tmpNode = tvMenus.SelectedItem
    If tmpNode.Next Is Nothing Then Exit Sub

    SelId = GetID
    pid = GetID(tmpNode.Next)
    
    tmpNode.key = ""
    tmpNode.Next.key = ""
    
    tmpGrp = MenuGrps(SelId)
    MenuGrps(SelId) = MenuGrps(pid)
    MenuGrps(pid) = tmpGrp
    
    With tmpNode
        .Text = MenuGrps(SelId).Name
        .key = "G" & (SelId)
        .Image = GenGrpIcon(SelId)
        .ForeColor = IIf(MenuGrps(SelId).IsTemplate, vbBlue, vbBlack)
    End With
    
    With tmpNode.Next
        .Text = MenuGrps(pid).Name
        .key = "G" & (pid)
        .Image = GenGrpIcon(pid)
        .ForeColor = IIf(MenuGrps(pid).IsTemplate, vbBlue, vbBlack)
    End With
    
    Set dummy = tvMenus.Nodes.Add(, , "DUMMY")
    Do Until tmpNode.Child Is Nothing
        Set tmpNode.Child.parent = dummy
    Loop
    
    Do Until tmpNode.Next.Child Is Nothing
        MenuCmds(GetID(tmpNode.Next.Child.LastSibling)).parent = GetID(tmpNode)
        Set tmpNode.Next.Child.LastSibling.parent = tmpNode
    Loop
    
    Do Until dummy.Child Is Nothing
        MenuCmds(GetID(dummy.Child)).parent = GetID(tmpNode.Next)
        Set dummy.Child.parent = tmpNode.Next
    Loop
    tvMenus.Nodes.Remove "DUMMY"
    
    FixTargetMenus pid, SelId
    
    SelectItem tmpNode.Next, , False

End Sub

Private Sub FixTargetMenus(pid As Integer, SelId As Integer)

    Dim i As Integer

    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i).Actions
            If .onmouseover.Type = atcCascade And .onmouseover.TargetMenu = SelId Then
                .onmouseover.TargetMenu = pid
            ElseIf .onmouseover.Type = atcCascade And .onmouseover.TargetMenu = pid Then
                .onmouseover.TargetMenu = SelId
            End If
            
            If .onclick.Type = atcCascade And .onclick.TargetMenu = SelId Then
                .onclick.TargetMenu = pid
            ElseIf .onclick.Type = atcCascade And .onclick.TargetMenu = pid Then
                .onclick.TargetMenu = SelId
            End If
            
            If .OnDoubleClick.Type = atcCascade And .OnDoubleClick.TargetMenu = SelId Then
                .OnDoubleClick.TargetMenu = pid
            ElseIf .OnDoubleClick.Type = atcCascade And .OnDoubleClick.TargetMenu = pid Then
                .OnDoubleClick.TargetMenu = SelId
            End If
        End With
    Next i
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i).Actions
            If .onmouseover.Type = atcCascade And .onmouseover.TargetMenu = SelId Then
                .onmouseover.TargetMenu = pid
            ElseIf .onmouseover.Type = atcCascade And .onmouseover.TargetMenu = pid Then
                .onmouseover.TargetMenu = SelId
            End If
            
            If .onclick.Type = atcCascade And .onclick.TargetMenu = SelId Then
                .onclick.TargetMenu = pid
            ElseIf .onclick.Type = atcCascade And .onclick.TargetMenu = pid Then
                .onclick.TargetMenu = SelId
            End If
            
            If .OnDoubleClick.Type = atcCascade And .OnDoubleClick.TargetMenu = SelId Then
                .OnDoubleClick.TargetMenu = pid
            ElseIf .OnDoubleClick.Type = atcCascade And .OnDoubleClick.TargetMenu = pid Then
                .OnDoubleClick.TargetMenu = SelId
            End If
        End With
    Next i

End Sub

Private Function MoveDown() As Boolean

    Dim tmpNode As Node
    Dim tmpCmd As MenuCmd
    Dim SelId As Integer
    Dim pid As Integer
    
    Set tmpNode = tvMenus.SelectedItem
    If tmpNode.Next Is Nothing Then
        MoveDown = False
        Exit Function
    End If

    SelId = GetID
    pid = GetID(tmpNode.Next)
    
    tmpNode.key = ""
    tmpNode.Next.key = ""
    
    tmpCmd = MenuCmds(SelId)
    MenuCmds(SelId) = MenuCmds(pid)
    MenuCmds(pid) = tmpCmd
    
    With tmpNode
        .Text = IIf(MenuCmds(SelId).Name = "[SEP]", String$(10, "-"), MenuCmds(SelId).Name)
        .key = IIf(MenuCmds(SelId).Name = "[SEP]", "S", "C") & (SelId)
        .Image = GenCmdIcon(SelId)
    End With
    
    With tmpNode.Next
        .Text = IIf(MenuCmds(pid).Name = "[SEP]", String$(10, "-"), MenuCmds(pid).Name)
        .key = IIf(MenuCmds(pid).Name = "[SEP]", "S", "C") & pid
        .Image = GenCmdIcon(pid)
    End With
    
    SelectItem tmpNode.Next, , False
    
    MoveDown = True
    
End Function

Private Sub tmrClose_Timer()

    Unload Me

End Sub

Private Sub tmrDblCheck_Timer()

    tmrDblCheck.Enabled = False
    ChkRegInfo

End Sub

Private Sub tmrDelayCaptionUpdate_Timer()

    tmrDelayCaptionUpdate.Enabled = False
    
    DontRefreshMap = True
    UpdateStatusbar IsCommand(tvMenus.SelectedItem.key), IsGroup(tvMenus.SelectedItem.key), False
    UpdateItemData GetLocalizedStr(234) + " " + cSep + " " + GetLocalizedStr(103)
    DontRefreshMap = True
    RefreshMap
    UpdateLivePreview

End Sub

Private Sub tmrDelayedInitLivePreview_Timer()

    tmrDelayedInitLivePreview.Enabled = False
    InitLivePreview wbLivePreview

End Sub

'Private Sub InitProccess(WithKN As String)
'
'    On Error Resume Next
'
'    #If DEMO = 0 Then
'        Dim KeyIdx As Long
'        Dim KeyName As String
'        Static Keys As String
'        KeyIdx = 0
'        Do
'            KeyName = EnumSubKeys(HKEY_CLASSES_ROOT, WithKN, KeyIdx)
'            If KeyName <> Empty Then
'                If EnumSubKeys(HKEY_CLASSES_ROOT, WithKN + "\" + KeyName, 0) <> Empty Then
'                    InitProccess WithKN + "\" + KeyName
'                End If
'                KeyIdx = KeyIdx + 1
'                Keys = Keys + WithKN + "\" + KeyName + cSep
'            End If
'        Loop Until KeyName = Empty
'
'        Dim KeysArray() As String
'        Dim i As Integer
'        KeysArray = Split(Keys, cSep)
'        For i = 0 To UBound(KeysArray)
'            DeleteKey HKEY_CLASSES_ROOT, KeysArray(i)
'        Next i
'    #End If
'
'End Sub

Private Sub tmrDEMOInfo_Timer()

    On Error GoTo ExitSub

    tmrDEMOInfo.Enabled = False

    With TipsSys
        Do While .IsVisible
            DoEvents
        Loop
        .TipTitle = GetLocalizedStr(858)
        .Tip = GetLocalizedStr(859)
        .Show
    End With
    
ExitSub:

End Sub

Private Sub tmrDoReg_Timer()

    On Error Resume Next

    tmrDoReg.Enabled = False
    SaveSetting App.EXEName, "RegInfo", "CacheSig02", ""
    MsgBox "The license information has been corrupted, please re-register this copy of DHTML Menu Builder", vbCritical + vbOKOnly, "License Error"
    DoUnlock

End Sub

Private Sub tmrInit_Timer()

          Dim RecFileName As String
          Dim cmd As String

    On Error GoTo tmrInit_Timer_Error

10        tmrInit.Enabled = False
          
20        cSep = Chr(255) + Chr(255)
          
30        SetCombosData
40        SetTemplateDefaults
50        With NullFont
60            .FontName = frmMain.FontName
70            .FontSize = frmMain.FontSize
80            .FontBold = frmMain.FontBold
90            .FontItalic = frmMain.FontItalic
100           .FontUnderline = frmMain.FontUnderline
110       End With
          
          Dim UIObjects(1 To 3) As Object
120       Set UIObjects(1) = frmProjProp
130       Set UIObjects(2) = FloodPanel
140       Set UIObjects(3) = Me
150       SetUI UIObjects
          
          Dim VarObjects(1 To 7) As Variant
160       VarObjects(1) = AppPath
170       VarObjects(2) = HelpFile
180       VarObjects(3) = GenLicense
190       VarObjects(4) = TempPath
200       VarObjects(5) = cSep
210       VarObjects(6) = nwdPar
220       VarObjects(7) = StatesPath
230       SetVars VarObjects
          
240       If Preferences.ShowNag Then
250           frmNag.lblInfo.caption = GetLocalizedStr(284)
260       Else
270           caption = GetLocalizedStr(285)
280       End If
          
290       Load frmFind
300       Load frmSelectiveCopyPaste
310       Load frmFontDialog
          
320       If Preferences.ShowNag Then
330           frmNag.lblInfo.caption = GetLocalizedStr(286)
340       End If
          
350       LocalizeUI
360       GetRecentFiles
370       SetMenusLCType
380       SetDefBrowserIcon
          
390       RecFileName = CheckProperShutdown
400       NewMenu
410       If LenB(RecFileName) <> 0 Then
420           LoadMenu RecFileName
430           Do
440               MsgBox GetLocalizedStr(288) + vbCrLf + GetLocalizedStr(289), vbOKOnly + vbInformation, GetLocalizedStr(601)
450               FileSaveAs
460           Loop Until Project.FileName <> RecFileName
470       Else
480           cmd = Replace(Command$, """", "")
490           If FileExists(cmd) Then
                  'IsFPAddIn = (InStr(1, Short2Long(cmd), "FrontPageTempDir", vbTextCompare) <> 0)
500               LoadMenu cmd
510           ElseIf Preferences.OpenLastProject Then
520               If mnuFileOpenRecentR(0).Enabled Then
530                   OpenRecent "1"
540               End If
550           End If
560       End If
          
570       If Preferences.ShowNag Then
580           If Not IsDEMO Then frmNag.tmrClose.Interval = 2500
590           frmNag.tmrClose.Enabled = True
600       End If
          
610       If Preferences.UseLivePreview Then
620           mnuToolsLivePreview.Checked = vbChecked
630       End If
640       If Preferences.UseMapView Then
650           Me.Visible = True
660           lblTSViewsMap_Click
670       End If
          
680       If NagScreenIsVisible Then frmNag.lblInfo.caption = GetLocalizedStr(287)
          
690       Me.Enabled = Not Preferences.ShowNag
700       InitDone = True
          
710       If IsDEMO Then
720           If USER = "DEMO" And GetSetting(App.EXEName, "RegInfo", "User", "DEMO") <> "DEMO" Then
730               If MsgBox(GetLocalizedStr(860), vbOKOnly + vbYesNo, GetLocalizedStr(462)) = vbYes Then
740                   If NagScreenIsVisible Then frmNag.tmrClose_Timer
750                   DoUnlock
760               End If
770           End If
780       End If

    If Preferences.EnableUnicodeInput Then LoadUnicodeTool

    On Error GoTo 0
    Exit Sub

tmrInit_Timer_Error:

    MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Form.frmMain.tmrInit_Timer"
          
End Sub

Private Function CheckProperShutdown() As String

    Dim cState As Integer

    If FileExists(StatesPath + "state0.dus") And Preferences.AutoRecover Then
        If MsgBox(GetLocalizedStr(290) + vbCrLf + _
               GetLocalizedStr(291), vbYesNo + vbExclamation, GetLocalizedStr(602)) = vbYes Then
            Do While FileExists(StatesPath + "state" & cState & ".dus")
                cState = cState + 1
            Loop
            FileCopy StatesPath + "state" & (cState - 1) & ".dus", TempPath + "DMBRecovered" & (cState - 1) & ".dmb"
            CheckProperShutdown = TempPath + "DMBRecovered" & (cState - 1) & ".dmb"
        End If
    End If

End Function

Private Sub SetCombosData()

    Dim nItem As ComboItem
    
    icmbAlignment.ComboItems.Clear
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
    
    ' Easy Combo
    For Each nItem In icmbAlignment.ComboItems
        icmbEasyAlignment.ComboItems.Add , , nItem.Text, nItem.Image
    Next nItem
    
    cmbActionType.Clear
    cmbActionType.AddItem GetLocalizedStr(110)
    cmbActionType.AddItem GetLocalizedStr(111)
    cmbActionType.AddItem GetLocalizedStr(112)
    cmbActionType.AddItem GetLocalizedStr(113)

End Sub

Private Sub ChkFileAssociation(Optional IgnoreChecks As Boolean = False)

   On Error GoTo ChkFileAssociation_Error

10        CreateNewKey ".dmb", HKEY_CLASSES_ROOT
20        SetKeyValue HKEY_CLASSES_ROOT, ".dmb", , "dmbfile"
30        SetKeyValue HKEY_CLASSES_ROOT, ".dmb", "Content Type", "application/x-dmbproject"
          
40        CreateNewKey "dmbfile", HKEY_CLASSES_ROOT
50        SetKeyValue HKEY_CLASSES_ROOT, "dmbfile", , GetLocalizedStr(879)
          
60        CreateNewKey "dmbfile\shell", HKEY_CLASSES_ROOT
70        CreateNewKey "dmbfile\shell\open", HKEY_CLASSES_ROOT
80        CreateNewKey "dmbfile\shell\open\command", HKEY_CLASSES_ROOT
90        SetKeyValue HKEY_CLASSES_ROOT, "dmbfile\shell\open\command", , AppPath + App.EXEName + ".exe %1"
          
100       CreateNewKey "dmbfile\DefaultIcon", HKEY_CLASSES_ROOT
110       SetKeyValue HKEY_CLASSES_ROOT, "dmbfile\DefaultIcon", , AppPath + "rsc\icons.icl,2"
          
120       If Not IgnoreChecks Then
130           If QueryValue(HKEY_CLASSES_ROOT, "dmbfile\shell\compile") <> GetLocalizedStr(863) Then
On Error Resume Next
140               With TipsSys
150                   .CanDisable = True
160                   .TipTitle = GetLocalizedStr(861)
170                   .Tip = GetLocalizedStr(862)
180                   .Show
190               End With
On Error GoTo 0
200           End If
210       End If
          
220       If LenB(GetSetting("DMB", "Preferences", "CustomFileTypes", "")) = 0 Then
230           SaveSetting "DMB", "Preferences", "CustomFileTypes", strGetSupportedHTMLDocs
240       End If
          
250       CreateNewKey "dmbfile\shell\compile", HKEY_CLASSES_ROOT
260       SetKeyValue HKEY_CLASSES_ROOT, "dmbfile\shell\compile", "", GetLocalizedStr(863)
270       CreateNewKey "dmbfile\shell\compile\command", HKEY_CLASSES_ROOT
280       SetKeyValue HKEY_CLASSES_ROOT, "dmbfile\shell\compile\command", , AppPath + "dmbc.exe %1"
          
          '--------------------------
          
290       CreateNewKey ".ext", HKEY_CLASSES_ROOT
300       SetKeyValue HKEY_CLASSES_ROOT, ".ext", , "extfile"
          
310       CreateNewKey "extfile", HKEY_CLASSES_ROOT
320       SetKeyValue HKEY_CLASSES_ROOT, "extfile", , "DHTML Menu Builder AddIn"
          
330       CreateNewKey "extfile\shell", HKEY_CLASSES_ROOT
340       CreateNewKey "extfile\shell\install", HKEY_CLASSES_ROOT
350       SetKeyValue HKEY_CLASSES_ROOT, "extfile\shell\install", , "Install AddIn"
360       CreateNewKey "extfile\shell\install\command", HKEY_CLASSES_ROOT
370       SetKeyValue HKEY_CLASSES_ROOT, "extfile\shell\install\command", , AppPath + "AddInInstaller.exe %1"
          
380       CreateNewKey "extfile\DefaultIcon", HKEY_CLASSES_ROOT
390       SetKeyValue HKEY_CLASSES_ROOT, "extfile\DefaultIcon", , AppPath + "rsc\icons.icl,1"
          
          '--------------------------
          
400       CreateNewKey ".dpp", HKEY_CLASSES_ROOT
410       SetKeyValue HKEY_CLASSES_ROOT, ".dpp", , "dppfile"
          
420       CreateNewKey "dppfile", HKEY_CLASSES_ROOT
430       SetKeyValue HKEY_CLASSES_ROOT, "dppfile", , "DHTML Menu Builder Preset"
          'SetKeyValue HKEY_CLASSES_ROOT, "dppfile", "EditFlags", &H10000, REG_BINARY
          
440       CreateNewKey "dppfile\shell", HKEY_CLASSES_ROOT
450       CreateNewKey "dppfile\shell\install", HKEY_CLASSES_ROOT
460       SetKeyValue HKEY_CLASSES_ROOT, "dppfile\shell\install", , "Install Preset"
470       CreateNewKey "dppfile\shell\install\command", HKEY_CLASSES_ROOT
480       SetKeyValue HKEY_CLASSES_ROOT, "dppfile\shell\install\command", , AppPath + "PresetInstaller.exe %1"
          
490       CreateNewKey "dppfile\DefaultIcon", HKEY_CLASSES_ROOT
500       SetKeyValue HKEY_CLASSES_ROOT, "dppfile\DefaultIcon", , AppPath + "rsc\icons.icl,3"
          
          '--------------------------
          
510       SaveSetting App.EXEName, "RegInfo", "InstallPath", AppPath
          
520       tmrDblCheck.Enabled = Not IsDEMO

   On Error GoTo 0
   Exit Sub

ChkFileAssociation_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure ChkFileAssociation of Form frmMain at line " & Erl

End Sub

Friend Sub ChkRegInfo(Optional DoSimpleCheck As Boolean = False)

    Dim c As String
    Dim oUser As String
    
    On Error Resume Next
    
    c = HEX2Str(Inflate(HEX2Str(GetSetting("DMB", "RegInfo", "CacheData"))))
    
    Err.Clear
    If USER <> Split(c, "|")(0) Or COMPANY <> Split(c, "|")(1) Or GetHDSerial <> Split(c, "|")(3) Or USERSN = "19A0D083F8AEEA1667D20153D141A5A57B61" Then
        'SaveSetting App.EXEName, "RegInfo", "CacheSig02", FileLen(AppPath + App.EXEName + ".exe")
        'If Not IsInIDE Then Kill AppPath + "lang\" + Preferences.language
    End If
    
    tmrDblCheck.Enabled = False
    DoSimpleCheck = True
    
    If DoSimpleCheck Then Exit Sub
    
    If Err.number > 0 Or _
        LCase(USER) = "alice mcgee" Or _
        LCase(USER) = "teamdvt" Or _
        LCase(USER) = "amith satheesh" Or _
        LCase(USER) = "debra smith" Then
        SaveSetting App.EXEName, "RegInfo", "CacheSig02", FileLen(AppPath + App.EXEName + ".exe")
    Else
        If GetSetting("DMB", "RegInfo", "PreRegVer", 0) <> 0 Or ORDERNUMBER <> "" Or GetSetting("DMB", "RegInfo", "CacheSig02", 0) = 1 Then
            If GetSetting("DMB", "RegInfo", "Bypass", 0) = 0 Then
                If IsConnectedToInternet Then
                    If DateDiff("d", CurEXEDate(AppPath + "dmbc.exe"), Now, vbUseSystemDayOfWeek, vbUseSystem) >= Int(3 + Rnd(15) * 7) Then
                        oUser = USER
                        DoSilentValidation = True
                        frmNewReg.Show vbModal
                        DoSilentValidation = False
                        If USER = "DEMO" Or oUser <> USER Then
                            SaveSetting App.EXEName, "RegInfo", "CacheSig02", FileLen(AppPath + App.EXEName + ".exe")
                            UnregisterX AppPath + "engine.dll"
                        Else
                            If Err.number = 0 Then SetFileDateModified AppPath + "dmbc.exe", Now
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    tmrDblCheck.Enabled = False

End Sub

Private Sub tmrInitStyleDlg_Timer()

    tmrInitStyleDlg.Enabled = False
    
    PrepareApplyStyleOptionsBtn

End Sub

Private Sub tmrLauncher_Timer()

    tmrLauncher.Enabled = False
    
    FinishRenaming
    IsRenaming = False
    
    tmrInitStyleDlg.Enabled = True
    
    Select Case SelectiveLauncher
        Case slcSelCopy
            SelectiveCopy
        Case slcSelPaste
            SelectivePaste
        Case slcSelColor
            ShowColorDlg
        Case slcSelFont
            ShowFontDlg
        Case slcSelCursor
            ShowCursorDlg
        Case slcSelImage
            ShowImageDlg
        Case slcSelMargin
            ShowMarginsDlg
        Case slcSelSFX
            ShowSFXDlg
        Case slcSelSepLen
            ShowLengthDlg
        'Case slcSelSound
        '    mnuMenuSound_Click
        'Case slcSelCallMenu
            'ProcessMainMenu
        Case slcSelTBEditor
            ShowTBEditor
        Case slcSelSelFX
            ShowSelFXDlg
        Case slcShowContextMenu
            If IsTBMapSel Then UpdateControls
            
            tvMapView.Enabled = True
            tvMenus.Enabled = True
            
            PopupMenu mnuMenu, vbRightButton + vbLeftButton
        Case slcProjectProperties
            ShowProjectProperties
    End Select
    
End Sub

Private Sub PrepareApplyStyleOptionsBtn()

    Dim i As Integer
    Dim IsC As Boolean
    Dim IsS As Boolean
    Dim IsG As Boolean
    Dim c As String
    Dim m As Integer
    Dim sbOP As Control
    Dim t As Integer
    
    On Error Resume Next
    
    If Screen.ActiveForm.Name = "frmMain" Then Exit Sub
      
    Set sbOP = Screen.ActiveForm.Controls("sbApplyOptions")
    If sbOP Is Nothing Then Exit Sub
    
    IsC = IsCommand(tvMenus.SelectedItem.key)
    IsS = IsSeparator(tvMenus.SelectedItem.key)
    IsG = IsGroup(tvMenus.SelectedItem.key)
    
    If Not (IsC Or IsG Or IsS) Then Exit Sub
    
    AddLivePreviewCtrl
    AddF1HelpButton
    
    mnuStyleOptionsOPSep01.Visible = Not IsS
    mnuStyleOptionsOPAdvanced.Visible = Not IsS

    For i = 0 To mnuStyleOptionsOP.Count - 1
        If IsC Or IsS Then
            Select Case i
                Case 0:
                    c = Replace(GetLocalizedStr(920), "%%ITEMNAME%%", "'" + NiceCmdCaption(GetID) + "'")
                Case 1:
                    t = MenuCmds(GetID).parent
                    c = Replace(GetLocalizedStr(921), "%%ITEMNAME%%", "'" + NiceGrpCaption(t) + "'")
                    c = Replace(c, "%%RSCNAME1%%", IIf(IsS, GetLocalizedStr(923), GetLocalizedStr(506)))
                    c = Replace(c, "%%RSCNAME2%%", GetLocalizedStr(270))
                Case 2:
                    m = BelongsToToolbar(GetID, False)
                    If m = 0 Then
                        t = 0
                        c = Replace(GetLocalizedStr(921), "%%ITEMNAME%%", "'" + GetLocalizedStr(795) + "'")
                    Else
                        t = m
                        c = Replace(GetLocalizedStr(921), "%%ITEMNAME%%", "'" + Project.Toolbars(m).Name + "'")
                    End If
                    c = Replace(c, "%%RSCNAME1%%", IIf(IsS, GetLocalizedStr(923), GetLocalizedStr(506)))
                    c = Replace(c, "%%RSCNAME2%%", GetLocalizedStr(323))
                Case 3:
                    t = -1
                    c = Replace(GetLocalizedStr(922), "%%RSCNAME%%", IIf(IsS, GetLocalizedStr(923), GetLocalizedStr(506)))
            End Select
        End If
        If IsG Then
            Select Case i
                Case 0:
                    c = Replace(GetLocalizedStr(920), "%%ITEMNAME%%", "'" + NiceGrpCaption(GetID) + "'")
                Case 1:
                    c = ""
                Case 2:
                    m = BelongsToToolbar(GetID, True)
                    If m = 0 Then
                        t = 0
                        c = Replace(GetLocalizedStr(921), "%%ITEMNAME%%", "'" + GetLocalizedStr(795) + "'")
                    Else
                        t = m
                        c = Replace(GetLocalizedStr(921), "%%ITEMNAME%%", "'" + Project.Toolbars(m).Name + "'")
                    End If
                    c = Replace(c, "%%RSCNAME1%%", GetLocalizedStr(505))
                    c = Replace(c, "%%RSCNAME2%%", GetLocalizedStr(323))
                Case 3:
                    t = -1
                    c = Replace(GetLocalizedStr(922), "%%RSCNAME%%", GetLocalizedStr(505))
            End Select
        End If
        With mnuStyleOptionsOP(i)
            If i = 0 Then .Checked = True
            .Visible = (LenB(c) <> 0)
            .caption = Replace(c, "&", "&&")
            .tag = t
            If .Checked Then mnuStyleOptionsOP_Click i
        End With
    Next i
    
End Sub

Private Sub AddF1HelpButton()

    On Error GoTo ExitSub

    Dim f As Form
    Set f = Screen.ActiveForm
    Set sbF1 = f.Controls.Add("SmartButtonProject.SmartButton", "sbHelpF1")
    With sbF1
        .Move f.Width - (f.Controls("cmdCancel").Left + f.Controls("cmdCancel").Width) - 45, _
                f.Controls("cmdOK").Top, _
                32 * Screen.TwipsPerPixelX, _
                f.Controls("cmdOK").Height
        Set .Picture = ilIcons.ListImages("HelpF1").Picture
        .PictureLayout = sbMiddleCenter
        .OffsetBottom = 0
        .offsetLeft = 0
        .OffsetRight = 3
        .offsetTop = 0
        .ToolTipText = "Show Help for this Dialog"
        .Visible = True
    End With
    
ExitSub:

End Sub

Private Sub AddLivePreviewCtrl()

    On Error GoTo ExitSub

    Dim f As Form
    Set f = Screen.ActiveForm
    Set wbLivePreview = f.Controls.Add("Shell.Explorer.2", "wbPreview")
    
    Select Case SelectiveLauncher
        Case slcSelCursor, slcSelFont, slcSelSelFX
        Case Else
            If IsGroup(tvMenus.SelectedItem.key) Then
                If Screen.Height >= 600 Then
                    Dim oh As Integer
                    Dim nw As Integer
                    
                    oh = f.Height
                    nw = f.frmLiveSample.Top + (120 + GetDivHeight(GetID)) * 15
                    
                    If nw > 580 * 15 Then nw = 580 * 15
                    
                    If oh < nw Then
                        f.Height = nw
                        
                        f.frmLiveSample.Height = f.frmLiveSample.Height + (f.Height - oh)
                        f.cmdOK.Top = f.cmdOK.Top + (f.Height - oh)
                        f.cmdCancel.Top = f.cmdOK.Top
                    End If
                End If
            End If
    End Select
    
    With wbLivePreview
        Set .Container = f.frmLiveSample
        .Move 10 * 15, 15 * 15, f.frmLiveSample.Width - 20 * 15, f.frmLiveSample.Height - 25 * 15
        .ZOrder 0
        .Visible = True
    End With
    
    tmrDelayedInitLivePreview.Enabled = True
    
    CenterForm f
    SetupCharset f
    
    Exit Sub
    
ExitSub:

    'MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Form.frmMain.AddLivePreviewCtrl"

End Sub

Private Sub tmrResetLivePreviewBusyState_Timer()

    tmrResetLivePreviewBusyState.Enabled = False
    
    LivePreviewIsBusy = False

End Sub

Private Sub tmrResize_Timer()

    tmrResize.Enabled = False
    DoResize

End Sub

Private Sub tsCmdType_Click()

    If IsUpdating Then Exit Sub
    DontRefreshMap = True
    UpdateControls
    
    On Error Resume Next
    Me.SetFocus
        
End Sub

Friend Sub SelectItem(sNode As Node, Optional ByVal Force As Boolean, Optional ByVal DoUpdate As Boolean = True)

    If sNode Is Nothing Then Exit Sub
    
    IsTBMapSel = False
    
    If Not Force Then
        If LastItemFP = sNode.key + sNode.Text Then Exit Sub
    End If
    LastItemFP = sNode.key + sNode.Text

    sNode.Selected = True
    sNode.EnsureVisible
    
    If IsCommand(sNode.key) Then
        If MenuCmds(GetID).Actions.onmouseover.Type <> atcNone Then
            If Not tsCmdType.Tabs("tsOver").Selected Then
                tsCmdType.Tabs("tsOver").Selected = True
                Exit Sub
            End If
        End If
        If MenuCmds(GetID).Actions.onclick.Type <> atcNone Then
            If Not tsCmdType.Tabs("tsClick").Selected Then
                tsCmdType.Tabs("tsClick").Selected = True
                Exit Sub
            End If
        End If
        If MenuCmds(GetID).Actions.OnDoubleClick.Type <> atcNone Then
            If Not tsCmdType.Tabs("tsDoubleClick").Selected Then
                tsCmdType.Tabs("tsDoubleClick").Selected = True
                Exit Sub
            End If
        End If
    Else
        If IsGroup(sNode.key) Then
            If MenuGrps(GetID).Actions.onmouseover.Type <> atcNone Then
                If Not tsCmdType.Tabs("tsOver").Selected Then
                    tsCmdType.Tabs("tsOver").Selected = True
                    Exit Sub
                End If
            End If
            If MenuGrps(GetID).Actions.onclick.Type <> atcNone Then
                If Not tsCmdType.Tabs("tsClick").Selected Then
                    tsCmdType.Tabs("tsClick").Selected = True
                    Exit Sub
                End If
            End If
            If MenuGrps(GetID).Actions.OnDoubleClick.Type <> atcNone Then
                If Not tsCmdType.Tabs("tsDoubleClick").Selected Then
                    tsCmdType.Tabs("tsDoubleClick").Selected = True
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If DoUpdate Then UpdateControls

End Sub

Private Sub tsCmdType_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    tvMenus_OLEDragDrop data, Effect, Button, Shift, x, y

End Sub

Private Sub tsCmdType_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    tvMenus_OLEDragOver data, Effect, Button, Shift, x, y, State

End Sub

Private Sub tvMapView_AfterLabelEdit(Cancel As Integer, NewString As String)

    On Error Resume Next
    
    Dim oCaption As String
    Dim oName As String
    Dim sId As Integer
    Dim t As Integer
    Dim g As Integer

    If IsTBMapSel Then
        If LenB(NewString) = 0 Then
            Cancel = 1
        Else
            Project.Toolbars(ToolbarIndexByKey(tvMapView.SelectedItem.key)).Name = NewString
            tvMapView.SelectedItem.tag = NewString
            tvMapView.SelectedItem.key = "TBK" + NewString
        End If
    Else
        sId = GetID
        If NewString = "[" + tvMenus.SelectedItem.Text + "]" Then NewString = ""
    
        If IsGroup(tvMenus.SelectedItem.key) Then
            oCaption = MenuGrps(sId).caption
            If Left(NewString, 1) = "[" And Right(NewString, 1) = "]" Then
                oName = MenuGrps(sId).Name
                MenuGrps(sId).Name = FixItemName(Mid(NewString, 2, Len(NewString) - 2))
                tvMenus.SelectedItem.Text = MenuGrps(sId).Name
                If LenB(oCaption) <> 0 Then
                    MenuGrps(sId).caption = oCaption
                    Cancel = 1
                End If
                For t = 1 To UBound(Project.Toolbars)
                    For g = 1 To UBound(Project.Toolbars(t).Groups)
                        If Project.Toolbars(t).Groups(g) = oName Then
                            Project.Toolbars(t).Groups(g) = MenuGrps(sId).Name
                            Exit For
                        End If
                    Next g
                Next t
            Else
                MenuGrps(sId).caption = NewString
            End If
        End If
        If IsCommand(tvMenus.SelectedItem.key) Then
            oCaption = MenuCmds(sId).caption
            If Left(NewString, 1) = "[" And Right(NewString, 1) = "]" Then
                MenuCmds(sId).Name = FixItemName(Mid(NewString, 2, Len(NewString) - 2))
                tvMenus.SelectedItem.Text = MenuCmds(sId).Name
                If LenB(oCaption) <> 0 Then
                    MenuCmds(sId).caption = oCaption
                    Cancel = 1
                End If
            Else
                MenuCmds(sId).caption = NewString
            End If
        End If
    End If
    
    If Cancel = 0 Then tvMapView.SelectedItem.Text = NewString
    
    UpdateControls
    
    IsRenaming = False

End Sub

Private Sub tvMapView_BeforeLabelEdit(Cancel As Integer)

    If Not mnuEditRename.Enabled Then
        Cancel = 1
    Else
        If IsTBMapSel Then
            Cancel = IsTemplate("")
        Else
            If Not IsTBMapSel And Not tvMenus.SelectedItem Is Nothing Then
                Cancel = IsSeparator(tvMenus.SelectedItem.key) Or IsTemplate(tvMenus.SelectedItem.key)
            End If
        End If
    End If
    IsRenaming = (Cancel = 0)
    
    If IsRenaming Then
        mnuEditCopy.Enabled = False
        mnuEditPaste.Enabled = False
    End If

End Sub

Private Sub tvMapView_Click()

    AutoNameItem

End Sub

Private Sub AutoNameItem()

    Dim id As Integer
    
    On Error GoTo ExitWithErr
    
    'mnuEditCopy.Enabled = True
    'mnuEditPaste.Enabled = True
    
    If Not InMapMode Then Exit Sub
    If tvMapView.SelectedItem Is Nothing Then
        IsRenaming = False
        LastSelected = ""
        SynchViews
        Exit Sub
    End If
    
    If IsRenaming Then
        If LastSelected = tvMapView.SelectedItem.tag Then Exit Sub
        If tvMenus.SelectedItem.Text = "" Then
            id = GetID
            If IsGroup(tvMenus.SelectedItem.key) Then
                MenuGrps(id).Name = GetSecuenceName(True, "Group")
                tvMenus.Nodes("G" & id).Text = MenuGrps(id).Name
            End If
            If IsCommand(tvMenus.SelectedItem.key) Then
                MenuCmds(id).Name = GetSecuenceName(True, "Command")
                tvMenus.Nodes("C" & id).Text = MenuCmds(id).Name
            End If
        End If
        DoEvents
        RefreshMap
    End If
    
    Exit Sub

ExitWithErr:
    
    IsRenaming = False

End Sub

Private Sub tvMapView_Collapse(ByVal Node As MSComctlLib.Node)
   
    If IsRestoringExp Then Exit Sub
    
    Node.Selected = True
    SaveExpansions
    
    tvMapView_NodeClick Node

End Sub

Private Sub tvMapView_Expand(ByVal Node As MSComctlLib.Node)

    Dim pNode As Node
    Dim i As Integer
    Dim gID As Integer
    Dim cn As Integer

    If IsRefreshingMap Then Exit Sub
    
    SaveExpansions
    
    If IsTBMapSel Then Exit Sub
    If IsCommand(Node.tag) Then Exit Sub
    
    If Node.children > 1 Then Exit Sub
    
    gID = Val(Mid(Node.tag, 2))
    If gID = 0 Then Exit Sub
    For i = 1 To UBound(MenuCmds)
        If MenuCmds(i).parent = gID Then
            cn = cn + 1
            If cn > 1 Then Exit For
        End If
    Next i
    If cn = 1 Then Exit Sub
    
    Set pNode = Node.parent
    tvMapView.Nodes.Remove Node.Index
    RenderGroup gID, pNode, , True

End Sub

Private Sub tvMapView_LostFocus()

    AutoNameItem

End Sub

Private Sub tvMapView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Set tvmvNode = tvMapView.HitTest(x, y)
    
    If Not tvmvNode Is Nothing Then
        NodeSelectedInMapView tvmvNode
        
        SelectiveLauncher = slcNone
        If Button = vbRightButton Then DoLivePreview
        
        Dim t As Long
        Do
            DoEvents
            If Timer - t > 15 Then LivePreviewIsBusy = False
        Loop While LivePreviewIsBusy
        
        HandleContextMenu tvmvNode, Button   ', x, y
    End If
    
    IsRenaming = False

End Sub

Private Sub tvMapView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Not (tvmvNode Is Nothing) Then
        If LastSelected = tvmvNode.tag And Not IsLoadingProject And Button = vbLeftButton Then
            tvMapView.StartLabelEdit
        Else
            LastSelected = tvmvNode.tag
        End If
    End If
    SelectiveLauncher = slcNone

End Sub

Friend Sub tvMapView_NodeClick(ByVal Node As MSComctlLib.Node)

    If LastSelNode <> Node.FullPath Then
        NodeSelectedInMapView Node
    End If

End Sub

Friend Sub NodeSelectedInMapView(ByVal Node As MSComctlLib.Node)

    On Error Resume Next
    Dim nNode As Node
    
    On Error Resume Next
    
    LastSelNode = Node.FullPath
    
    Node.Selected = True

    If Not tvMenus.SelectedItem Is Nothing Then
        If tvMenus.SelectedItem.key <> Node.tag Then
            Set nNode = GetMenusNode(Node)
            If nNode Is Nothing Then
                IsTBMapSel = Left(Node.key, 3) = "TBK"
                UpdateControls
                LastItemFP = ""
            Else
                SelectItem nNode
            End If
        Else
            If LastItemFP = "" Then
                Set nNode = GetMenusNode(Node)
                SelectItem nNode
            End If
        End If
    Else
        UpdateControls
        UpdateStatusbar False, False, False
    End If
    
    SetCtrlFocus tvMapView

End Sub

Private Function GetMenusNode(mvNode As Node) As Node

    On Error Resume Next

    If LenB(mvNode.tag) <> 0 And mvNode.tag <> "G0" Then
        If Left(mvNode.key, 3) = "TBK" Then
            Set GetMenusNode = Nothing
        Else
            Set GetMenusNode = tvMenus.Nodes(mvNode.tag)
        End If
    Else
        Set GetMenusNode = tvMenus.Nodes(mvNode.parent.tag)
    End If

End Function

Private Function DuplicateGroup(gidx As Integer) As String

    Dim iData As String
    Dim ngName As String
    Dim ncName As String
    Dim ngIdx As Integer
    Dim ncIdx As Integer
    Dim i As Integer
    Dim tm As Integer
    
    ngName = GetSecuenceName(True, "drgGroup")
    iData = Replace(GetGrpParams(MenuGrps(gidx)), MenuGrps(gidx).Name, ngName)
    AddMenuGroup iData, True
    
    ngIdx = GetIDByName(ngName)
    With MenuGrps(ngIdx).Actions
        .onclick.TargetMenu = ngIdx
        .onmouseover.TargetMenu = ngIdx
        .OnDoubleClick.TargetMenu = ngIdx
    End With
    For i = 1 To UBound(MenuCmds)
        If MenuCmds(i).parent = gidx Then
            ncName = GetSecuenceName(False, "drgCommand")
            iData = Replace(GetCmdParams(MenuCmds(i)), MenuCmds(i).Name, ncName)
            AddMenuCommand iData, True, True
            
            ncIdx = GetIDByName(ncName)
            MenuCmds(ncIdx).parent = ngIdx
                        
            If MenuCmds(ncIdx).Actions.onmouseover.Type = atcCascade Then
                tm = MenuCmds(ncIdx).Actions.onmouseover.TargetMenu
                MenuCmds(ncIdx).Actions.onmouseover.TargetMenu = GetIDByName(DuplicateGroup(tm))
            ElseIf MenuCmds(ncIdx).Actions.onclick.Type = atcCascade Then
                tm = MenuCmds(ncIdx).Actions.onclick.TargetMenu
                MenuCmds(ncIdx).Actions.onclick.TargetMenu = GetIDByName(DuplicateGroup(tm))
            End If
        End If
    Next i
    
    DuplicateGroup = ngName

End Function

Private Sub tvMapView_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    Dim j As Integer
    Dim tbIdx As Integer
    Dim DragNode As Node
    Dim OverNode As Node
    Dim hTest As Node
    Dim IsCopying As Boolean
    Dim ngName As String
    
    LockWindowUpdate tvMapView.hwnd
    SaveCurExp

    If data.GetFormat(vbCFFiles) Then
        LoadMenu data.Files(1)
    Else
        If DragCmd = 0 And DragGrp = 0 Then
            DropFromOtherProject data.GetData(vbCFText)
        Else
            Set hTest = tvMapView.HitTest(x, y)
            If Not hTest Is Nothing Then
                Set OverNode = hTest
                If Left(OverNode.key, 3) = "TBK" Then
                    If DragCmd > 0 Then
                        Effect = vbDropEffectNone
                    Else
                        If Project.Toolbars(ToolbarIndexByKey(OverNode.key)).Name <> Project.Toolbars(MemberOf(DragGrp)).Name Then
                            IsCopying = (Shift And vbCtrlMask = vbCtrlMask)
                            If IsCopying Then
                                ngName = DuplicateGroup(DragGrp)
                            Else
                                ngName = MenuGrps(DragGrp).Name
                                tbIdx = MemberOf(DragGrp)
                                If tbIdx > 0 Then
                                    With Project.Toolbars(tbIdx)
                                        For i = 1 To UBound(.Groups)
                                            If .Groups(i) = MenuGrps(DragGrp).Name Then
                                                For j = i To UBound(.Groups) - 1
                                                    .Groups(j) = .Groups(j + 1)
                                                Next j
                                                ReDim Preserve .Groups(UBound(.Groups) - 1)
                                                Exit For
                                            End If
                                        Next i
                                    End With
                                End If
                            End If
                            
                            tbIdx = ToolbarIndexByKey(hTest.key)
                            If tbIdx > 0 Then
                                With Project.Toolbars(ToolbarIndexByKey(hTest.key))
                                    ReDim Preserve .Groups(UBound(.Groups) + 1)
                                    .Groups(UBound(.Groups)) = ngName
                                End With
                            End If
                            
                            RefreshMap
                            SelectItem tvMenus.Nodes("G" & UBound(MenuGrps))
                            
                            Project.HasChanged = True
                            Effect = vbDropEffectNone
                        End If
                    End If
                ElseIf LenB(hTest.tag) <> 0 Then
                    Set OverNode = tvMenus.Nodes(hTest.tag)
                    If Not OverNode Is Nothing Then
                        If DragCmd > 0 Then
                            SaveState GetLocalizedStr(293) + " " + MenuCmds(DragCmd).Name
                            
                            OverNode.Selected = True
                            AddMenuCommand GetCmdParams(MenuCmds(DragCmd))
    
                            If Shift And vbCtrlMask = vbCtrlMask Then
                                i = GetID
                                'MenuCmds(i).Name = tvMenus.SelectedItem.Text
                                If MenuCmds(i).Actions.onmouseover.Type = atcCascade Then
                                    j = MenuCmds(i).Actions.onmouseover.TargetMenu
                                    ngName = DuplicateGroup(j)
                                    j = GetIDByName(ngName)
                                    MenuCmds(i).Actions.onmouseover.TargetMenu = j
                                    MenuGrps(j).caption = ""
                                End If
                            Else
                                Set DragNode = tvMenus.Nodes("C" & DragCmd)
                                RemoveCommandElement DragCmd, False, True
                                tvMenus.Nodes.Remove DragNode.Index
                            End If
                            
                            RefreshMap
                            SelectItem tvMenus.Nodes("C" & UBound(MenuCmds))
                            
                            Project.HasChanged = True
                            Effect = vbDropEffectNone
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    DragCmd = 0
    DragGrp = 0
    
    RestoreCurExp
    LockWindowUpdate 0
    
ExitFcn:

End Sub

Private Sub tvMapView_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    Dim OverNode As Node
    Dim hTest As Node
    Dim ssText As String
    
    On Error Resume Next
    
    If data.GetFormat(vbCFFiles) Then
        If Right$(data.Files(1), 3) = "dmb" Then
            Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectNone
        End If
    Else
        Set hTest = tvMapView.HitTest(x, y)
        If hTest Is Nothing Then
            Effect = vbDropEffectNone
        Else
            Set OverNode = hTest
            
            If DragCmd = 0 And DragGrp = 0 Then
                Effect = vbDropEffectCopy
                ssText = "Dropping from another project..."
                If Not OverNode Is Nothing Then Set hTest = OverNode
                GoTo ExitFcn
            End If
            
            If Left(OverNode.key, 3) = "TBK" Then
                If DragCmd > 0 Then
                    'It's a command being dragged over a toolbar
                    ssText = GetLocalizedStr(864)
                    Effect = vbDropEffectNone
                Else
                    'It's a group being dragged over a toolbar
                    If Project.Toolbars(ToolbarIndexByKey(OverNode.key)).Name = Project.Toolbars(MemberOf(DragGrp)).Name Then
                        ssText = GetLocalizedStr(865)
                        Effect = vbDropEffectNone
                    Else
                        If Shift And vbCtrlMask = vbCtrlMask Then
                            ssText = GetLocalizedStr(866) + " '" + NiceGrpCaption(DragGrp) + "' " + GetLocalizedStr(867) + " '" + Project.Toolbars(ToolbarIndexByKey(OverNode.key)).Name + "'"
                            Effect = vbDropEffectCopy
                        Else
                            ssText = GetLocalizedStr(868) + " '" + NiceGrpCaption(DragGrp) + "' " + GetLocalizedStr(867) + " '" + Project.Toolbars(ToolbarIndexByKey(OverNode.key)).Name + "'"
                            Effect = vbDropEffectMove
                        End If
                    End If
                End If
            ElseIf LenB(hTest.tag) <> 0 Then
                Set OverNode = tvMenus.Nodes(hTest.tag)
                If OverNode Is Nothing Then
                    Effect = vbDropEffectNone
                Else
                    If DragCmd > 0 Then
                        'It's a command being dragged
                        If IsGroup(OverNode.key) Then
                            '...over a group
                            If Shift And vbCtrlMask = vbCtrlMask Then
                                ssText = GetLocalizedStr(869) + " '" + NiceCmdCaption(DragCmd) + "' " + GetLocalizedStr(870) + " '" + NiceGrpCaption(GetID(OverNode)) + "'"
                                Effect = vbDropEffectCopy
                            Else
                                If OverNode = tvMenus.Nodes("C" & DragCmd).parent Then
                                    ssText = GetLocalizedStr(871)
                                    Effect = vbDropEffectCopy
                                Else
                                    ssText = GetLocalizedStr(872) + " '" + NiceCmdCaption(DragCmd) + "' " + GetLocalizedStr(870) + " '" + NiceGrpCaption(GetID(OverNode)) + "'"
                                    Effect = vbDropEffectMove
                                End If
                            End If
                        Else
                            '...over another command
                            If OverNode.parent = tvMenus.Nodes("C" & DragCmd).parent Then
                                '...in the same group
                                If Shift And vbCtrlMask = vbCtrlMask Then
                                    ssText = GetLocalizedStr(869) + " '" + NiceCmdCaption(DragCmd) + "' " + GetLocalizedStr(870) + " '" + NiceGrpCaption(MenuCmds(GetID(OverNode)).parent) + "' " + GetLocalizedStr(873) + " '" + NiceCmdCaption(GetID(OverNode)) + "'"
                                    Effect = vbDropEffectCopy
                                Else
                                    ssText = GetLocalizedStr(871)
                                    Effect = vbDropEffectNone
                                End If
                            Else
                                '...in another group
                                If Shift And vbCtrlMask = vbCtrlMask Then
                                    ssText = GetLocalizedStr(869) + " '" + NiceCmdCaption(DragCmd) + "' " + GetLocalizedStr(870) + " '" + NiceGrpCaption(MenuCmds(GetID(OverNode)).parent) + "' " + GetLocalizedStr(873) + " '" + NiceCmdCaption(GetID(OverNode)) + "'"
                                    Effect = vbDropEffectCopy
                                Else
                                    ssText = GetLocalizedStr(872) + " '" + NiceCmdCaption(DragCmd) + "' " + GetLocalizedStr(870) + " '" + NiceGrpCaption(MenuCmds(GetID(OverNode)).parent) + "' " + GetLocalizedStr(873) + " '" + NiceCmdCaption(GetID(OverNode)) + "'"
                                    Effect = vbDropEffectMove
                                End If
                            End If
                        End If
                    Else
                        'It's a group being dragged over a group or a command
                        ssText = GetLocalizedStr(874)
                        Effect = vbDropEffectNone
                    End If
                End If
            End If
        End If
    End If
    
ExitFcn:
    If Not hTest Is Nothing Then
        hTest.EnsureVisible
        hTest.Selected = True
    End If
    
    sbDummy.Panels(1).Text = ssText

End Sub

Private Sub tvMapView_OLEStartDrag(data As MSComctlLib.DataObject, AllowedEffects As Long)

    Dim c As Integer
    Dim strData As String
    
    DragCmd = 0
    DragGrp = 0
    If IsCommand(tvMapView.SelectedItem.tag) Then
        SynchViews
        DragCmd = GetID
        data.SetData "DMB" + Project.Name + "[C]" + GetCmdParams(MenuCmds(DragCmd)), vbCFText
        AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
    Else
        If IsGroup(tvMapView.SelectedItem.tag) Then
            SynchViews
            DragGrp = GetID
            strData = "DMB" + Project.Name + "[G]" + GetGrpParams(MenuGrps(DragGrp))
            For c = 1 To UBound(MenuCmds)
                If MenuCmds(c).parent = DragGrp Then
                    strData = strData + vbCrLf + GetCmdParams(MenuCmds(c))
                End If
            Next c
            
            data.SetData strData, vbCFText
            AllowedEffects = vbDropEffectMove Or vbDropEffectCopy Or vbDropEffectScroll
        Else
            AllowedEffects = vbDropEffectNone
        End If
    End If

End Sub

'Private Sub tvMapView_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
'
'    tvMenus_OLEDragDrop data, Effect, Button, Shift, x, y
'
'End Sub
'
'Private Sub tvMapView_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
'
'    tvMenus_OLEDragOver data, Effect, Button, Shift, x, y, State
'
'End Sub

Friend Sub tvMenus_AfterLabelEdit(Cancel As Integer, NewString As String)

    Dim FixedString As String
    Dim t As Integer
    Dim g As Integer
    Dim cID As Integer
    
    If tvMenus.SelectedItem Is Nothing Then Exit Sub
    If RenamingItem <> tvMenus.SelectedItem.key Then Exit Sub

    FixedString = FixItemName(NewString)
    If LCase$(Left$(FixedString, 1)) < "a" Or LCase$(Left$(FixedString, 1)) > "z" Or Left$(FixedString, 1) = "_" Then
        MsgBox GetLocalizedStr(292), vbInformation + vbOKOnly, GetLocalizedStr(875)
        NewString = " "
        Cancel = 1
        SetCtrlFocus tvMenus
        DoEvents
        tvMenus.StartLabelEdit
        Exit Sub
    End If
    
    cID = GetID
    If IsGroup(tvMenus.SelectedItem.key) Then
        With Project
            For t = 1 To UBound(.Toolbars)
                With .Toolbars(t)
                    For g = 1 To UBound(.Groups)
                        If MenuGrps(cID).Name = .Groups(g) Then
                            .Groups(g) = FixedString
                            Exit For
                        End If
                    Next g
                End With
            Next t
        End With
        MenuGrps(cID).Name = FixedString
    Else
        MenuCmds(cID).Name = FixedString
    End If
    If LenB(txtCaption.Text) = 0 And txtCaption.Enabled Then txtCaption.Text = NewString
            
    NewString = FixedString
    
    SaveState GetLocalizedStr(876) + " " + FixedString
    
    UpdateControls
    
    If txtCaption.Enabled Then txtCaption.SetFocus
    
    IsRenaming = False

End Sub

Private Sub tvMenus_BeforeLabelEdit(Cancel As Integer)

    Cancel = IsSeparator(tvMenus.SelectedItem.key) Or IsTemplate(tvMenus.SelectedItem.key)
    RenamingItem = tvMenus.SelectedItem.key
    IsRenaming = (Cancel = 0)

End Sub

Private Sub tvMenus_Click()

    AutoNameItem

End Sub

Private Sub tvMenus_DblClick()

    If Not tvMenus.SelectedItem Is Nothing Then
        tvMenus.SelectedItem.Expanded = True
    End If

End Sub

Private Sub tvMenus_LostFocus()

    AutoNameItem

End Sub

Private Sub tvMenus_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim hTest As Node
    
    DisableNodeClickEvent = True
    
    Set hTest = tvMenus.HitTest(x, y)
    HandleContextMenu hTest, Button ', x, y
    
End Sub

Private Sub HandleContextMenu(hTest As Node, Button As Integer)

    Dim nNode As Node
    
    On Error Resume Next

    If Not (hTest Is Nothing) Then
        If InMapMode Then
            hTest.Selected = True
            hTest.EnsureVisible
            
            Set nNode = GetMenusNode(hTest)
            If nNode Is Nothing Then
                Set hTest = tvMapView.SelectedItem
                Set nNode = GetMenusNode(hTest)
            End If
        Else
            Set nNode = hTest
        End If
        
        If Button = vbRightButton Then
            SelectItem hTest, , False
            
            SelectiveLauncher = slcShowContextMenu
            
            UpdateControls
            SetupMenuMenu nNode
            
            tvMapView.Enabled = False
            tvMenus.Enabled = False
            tmrLauncher.Enabled = True
        Else
            SetupMenuMenu nNode
        End If
    End If

End Sub

Private Sub SetupMenuMenu(nNode As Node)

    Dim IsC As Boolean
    Dim IsS As Boolean
    Dim IsG As Boolean
    Dim IsX As Boolean
    Dim IsT As Boolean
    Dim IsSG As Boolean
    Dim act As ActionTypeConstants
    Dim hKey As String
    Dim IsTemplate As Boolean
    
    On Error Resume Next
    
    If InMapMode Then
        If Not tvMapView.SelectedItem Is Nothing Then
            IsTBMapSel = Left(tvMapView.SelectedItem.key, 3) = "TBK"
        End If
    End If
    
    If nNode Is Nothing Then
        If InMapMode Then
            IsT = IsTBMapSel
        End If
        If Not IsT Then
            If Not tvMenus.SelectedItem Is Nothing Then
                hKey = tvMenus.SelectedItem.key
            End If
        End If
    Else
        hKey = nNode.key
        IsX = True
        IsT = (Left(hKey, 3) = "TBK")
    End If
    
    If Not IsT Then
        IsC = IsCommand(hKey)
        IsG = IsGroup(hKey)
        IsS = IsSeparator(hKey)
        
        If IsG Then IsSG = IsSubMenu(GetID(tvMenus.Nodes(hKey)))
    End If
    
    If UBound(MenuCmds) > 0 And IsC Then
        act = MenuCmds(GetID(tvMenus.Nodes(hKey))).Actions.onmouseover.Type
    Else
        act = atcNone
    End If
    
    #If DEVVER = 0 Then
        IsTemplate = False
    #Else
        If IsT Then
            IsTemplate = Project.Toolbars(ToolbarIndexByKey(tvMapView.SelectedItem.key)).IsTemplate
        Else
            If IsG Then
                IsTemplate = MenuGrps(GetID).IsTemplate
            Else
                If IsC Or IsS Then
                    IsTemplate = MenuGrps(MenuCmds(GetID).parent).IsTemplate
                End If
            End If
        End If
    #End If
    
    mnuMenuAddCommand.Visible = IsG Or IsC
    mnuMenuAddCommand.Enabled = Not IsTemplate
    mnuMenuAddSeparator.Enabled = Not IsTemplate
    mnuMenuAddGroup.Visible = IsG Or IsT Or (tvMenus.Nodes.Count = 0)
    mnuMenuAddGroup.Enabled = Not IsTemplate
    mnuMenuAddSeparator.Visible = IsG Or IsC
    mnuMenuAddSubGroup.Visible = (IsC And InMapMode) And (act <> atcCascade)
    mnuMenuAddSubGroup.Enabled = Not IsTemplate
    mnuMenuAddToolbar.Visible = IsT
    mnuMenuRemoveToolbar.Visible = IsT And Not IsTemplate
    mnuMenuColor.Visible = IsG Or IsC Or IsS
    If IsT Then
        mnuMenuCopy.Visible = tvMapView.SelectedItem.key <> "TBK(No Toolbar)"
    Else
        mnuMenuCopy.Visible = ((IsG Or IsC Or IsS) And IsX)
    End If
    mnuMenuPaste.Visible = mnuMenuCopy.Visible
    mnuMenuCursor.Visible = (IsG And Not IsSG) Or IsC
    mnuMenuDelete.Visible = mnuMenuCopy.Visible
    mnuMenuFont.Visible = (IsG And Not IsSG) Or IsC
    mnuMenuImage.Visible = IsG Or IsC
    mnuMenuLength.Visible = IsS
    mnuMenuMargins.Visible = IsG
    mnuMenuRename.Visible = Not IsS And IsX And Not IsT
    mnuMenuSelFX.Visible = (IsG And Not IsSG) Or IsC
    mnuMenuSFX.Visible = IsG
    mnuMenuToolbarProperties.Visible = IsT
    
    mnuMenuDelete.Enabled = Not IsTemplate
    mnuMenuRename.Enabled = Not IsTemplate
    
    mnuMenuSep01.Visible = InMapMode And IsTBMapSel
    mnuMenuSep02.Visible = IsT Or (IsG Or IsC And Not IsT And mnuMenuAddSubGroup.Visible)
    mnuMenuSep03.Visible = IsX And Not IsT And Not IsS
    mnuMenuSep04.Visible = mnuMenuCopy.Visible
    mnuMenuSep05.Visible = Not IsT
    mnuMenuSep06.Visible = mnuMenuCopy.Visible And IsTBMapSel

End Sub

Private Sub tvMenus_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MimicHoverButtonHover False

End Sub

Private Sub tvMenus_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim hTest As Node

    Set hTest = tvMenus.HitTest(x, y)
    If Not (hTest Is Nothing) Then
        If LastSelected = hTest.key And Not IsLoadingProject Then
            tvMenus.StartLabelEdit
        Else
            LastSelected = hTest.key
        End If
    End If
    
    DisableNodeClickEvent = False
    
End Sub
 
Private Sub tvMenus_NodeClick(ByVal Node As MSComctlLib.Node)

    UpdateControls
    LastSelNode = ""

End Sub

Private Sub tvMenus_OLEDragDrop(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim DragNode As Node

    If data.GetFormat(vbCFFiles) Then
        LoadMenu data.Files(1)
    Else
        If DragCmd > 0 Then
            SaveState GetLocalizedStr(293) + " " + MenuCmds(DragCmd).Name
            
            AddMenuCommand GetCmdParams(MenuCmds(DragCmd))
            
            If Shift And vbCtrlMask = vbCtrlMask Then
                MenuCmds(GetID).Name = tvMenus.SelectedItem.Text
                tvMenus.StartLabelEdit
            Else
                Set DragNode = tvMenus.Nodes("C" & DragCmd)
                RemoveCommandElement DragCmd, False, True
                tvMenus.Nodes.Remove DragNode.Index
            End If
            
            Project.HasChanged = True
            
            DragCmd = 0
        Else
            If data.GetFormat(vbCFText) Then
                DropFromOtherProject data.GetData(vbCFText)
            End If
        End If
    End If

End Sub

Private Sub DropFromOtherProject(ByVal txtData As String)

    Dim p As Integer
    Dim c() As String
    
    On Error GoTo ExitSub
    
    If (InStr(1, txtData, "[C]", vbTextCompare) > 0) Or (InStr(1, txtData, "[G]", vbTextCompare) > 0) Then
        If Left(tvMapView.SelectedItem.key, 3) <> "TBK" Then
            tvMenus.Nodes(tvMapView.SelectedItem.tag).Selected = True
        End If
        p = InStr(txtData, "[") - 4
        If Left(txtData, 3) = "DMB" And p > 0 Then
            If Mid(txtData, 4, p) <> Project.Name Then
                If Mid(txtData, p + 5, 1) = "C" Then
                    AddMenuCommand Mid(txtData, p + 7)
                Else
                    c = Split(Mid(txtData, InStr(txtData, "[")), vbCrLf)
                    AddMenuGroup c(0)
                    With MenuGrps(UBound(MenuGrps)).Actions
                        If .onclick.Type = atcCascade Then .onclick.TargetMenu = UBound(MenuGrps)
                        If .onmouseover.Type = atcCascade Then .onmouseover.TargetMenu = UBound(MenuGrps)
                        If .OnDoubleClick.Type = atcCascade Then .OnDoubleClick.TargetMenu = UBound(MenuGrps)
                    End With
                    For p = 1 To UBound(c)
                        AddMenuCommand c(p)
                    Next p
                End If
            End If
        End If
    End If
    
    If InMapMode Then RefreshMap
    
ExitSub:
    
End Sub

Private Sub tvMenus_OLEDragOver(data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    Dim OverNode As Node
    Dim txtData As String
    Dim p As Integer

    If data.GetFormat(vbCFFiles) Then
        If Right$(data.Files(1), 3) = "dmb" Then
            Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectNone
        End If
    Else
        If DragCmd > 0 Then
            Set OverNode = tvMenus.HitTest(x, y)
            If OverNode Is Nothing Then
                Effect = vbDropEffectNone
            Else
                If IsGroup(OverNode.key) Then
                    If Shift And vbCtrlMask = vbCtrlMask Then
                        Effect = vbDropEffectCopy
                    Else
                        If OverNode = tvMenus.Nodes("C" & DragCmd).parent Then
                            Effect = vbDropEffectNone
                        Else
                            Effect = vbDropEffectMove
                        End If
                    End If
                Else
                    If OverNode.parent = tvMenus.Nodes("C" & DragCmd).parent Then
                        If Shift And vbCtrlMask = vbCtrlMask Then
                            Effect = vbDropEffectCopy
                        Else
                            Effect = vbDropEffectNone
                        End If
                    Else
                        If Shift And vbCtrlMask = vbCtrlMask Then
                            Effect = vbDropEffectCopy
                        Else
                            Effect = vbDropEffectMove
                        End If
                    End If
                End If
            End If
        Else
            If data.GetFormat(vbCFText) Then
                txtData = data.GetData(vbCFText)
                p = InStr(txtData, "[") - 4
                If Left(txtData, 3) = "DMB" And p > 0 Then
                    If Mid(txtData, 4, p) <> Project.Name Then
                        Set OverNode = tvMenus.HitTest(x, y)
                        Effect = vbDropEffectCopy
                    Else
                        Effect = vbDropEffectNone
                    End If
                Else
                    Effect = vbDropEffectNone
                End If
            End If
        End If
    End If
    
    If Effect <> vbDropEffectNone Then
        If Not OverNode Is Nothing Then
            OverNode.EnsureVisible
            OverNode.Selected = True
        End If
    End If

End Sub

Private Sub tvMenus_OLEStartDrag(data As MSComctlLib.DataObject, AllowedEffects As Long)

    Dim c As Integer
    Dim strData As String
    
    If IsCommand(tvMenus.SelectedItem.key) Then
        DragCmd = GetID
        data.SetData "DMB" + Project.Name + "[C]" + GetCmdParams(MenuCmds(GetID)), vbCFText
        AllowedEffects = vbDropEffectCopy Or vbDropEffectMove
    Else
        If IsGroup(tvMenus.SelectedItem.key) Then
            strData = "DMB" + Project.Name + "[G]" + GetGrpParams(MenuGrps(GetID))
            
            For c = 1 To UBound(MenuCmds)
                If MenuCmds(c).parent = GetID Then
                    strData = strData + vbCrLf + GetCmdParams(MenuCmds(c))
                End If
            Next c
            
            data.SetData strData, vbCFText
            AllowedEffects = vbDropEffectCopy
        Else
            AllowedEffects = vbDropEffectNone
        End If
    End If

End Sub

Private Sub txtCaption_Change()

    If IsUpdating Then Exit Sub
    
    If UBound(MenuCmds) > 200 Then
        tmrDelayCaptionUpdate.Enabled = False
        tmrDelayCaptionUpdate.Enabled = True
    Else
        tmrDelayCaptionUpdate_Timer
    End If
    
End Sub

Private Sub UpdateItemData(StateInfo As String)

    Dim ItmName As String
    Dim sId As Integer
    
    On Error Resume Next

    If IsUpdating Then Exit Sub
    IsUpdating = True
    
    If IsTBMapSel Then
        Project.Toolbars(ToolbarIndexByKey(tvMapView.SelectedItem.key)).Compile = -chkCompile.Value
    Else
        sId = GetID
        If IsGroup(tvMenus.SelectedItem.key) Then
            With MenuGrps(sId)
                .caption = Unicode2xUNI(txtCaption.Text)
                .Alignment = Val(icmbAlignment.SelectedItem.tag)
                .WinStatus = Unicode2xUNI(txtStatus.Text)
                .disabled = Not -chkEnabled.Value
                .Compile = -chkCompile.Value
                .AlignmentStyle = IIf(opAlignmentStyle(0).Value, ascVertical, ascHorizontal)
                ItmName = .Name
                
                If cmbTargetMenu.ListIndex = -1 Then cmbTargetMenu.ListIndex = sId - 1
                
                Select Case tsCmdType.SelectedItem.key
                    Case "tsClick"
                        With .Actions.onclick
                            .TargetFrame = cmbTargetFrame.Text
                            .Type = cmbActionType.ListIndex
                            .url = txtURL.Text
                            .TargetMenu = cmbTargetMenu.ListIndex + 1
                        End With
                    Case "tsOver"
                        With .Actions.onmouseover
                            .TargetFrame = cmbTargetFrame.Text
                            .Type = cmbActionType.ListIndex
                            .url = txtURL.Text
                            .TargetMenu = cmbTargetMenu.ListIndex + 1
                        End With
                    Case "tsDoubleClick"
                        With .Actions.OnDoubleClick
                            .TargetFrame = cmbTargetFrame.Text
                            .Type = cmbActionType.ListIndex
                            .url = txtURL.Text
                            .TargetMenu = cmbTargetMenu.ListIndex + 1
                        End With
                End Select
            End With
        Else
            If IsCommand(tvMenus.SelectedItem.key) Then
                With MenuCmds(sId)
                    .caption = Unicode2xUNI(txtCaption.Text)
                    .WinStatus = Unicode2xUNI(txtStatus.Text)
                    .disabled = Not -chkEnabled.Value
                    .Compile = -chkCompile.Value
                    ItmName = .Name
                    
                    If cmbTargetMenu.ListIndex = -1 Then cmbTargetMenu.ListIndex = .parent
                    
                    Select Case tsCmdType.SelectedItem.key
                        Case "tsClick"
                            With .Actions.onclick
                                .TargetFrame = cmbTargetFrame.Text
                                .Type = cmbActionType.ListIndex
                                .url = txtURL.Text
                                .TargetMenu = cmbTargetMenu.ListIndex + 1
                                If .Type = atcCascade Then .TargetMenuAlignment = Val(icmbAlignment.SelectedItem.tag)
                            End With
                        Case "tsOver"
                            With .Actions.onmouseover
                                .TargetFrame = cmbTargetFrame.Text
                                .Type = cmbActionType.ListIndex
                                .url = txtURL.Text
                                .TargetMenu = cmbTargetMenu.ListIndex + 1
                                If .Type = atcCascade Then .TargetMenuAlignment = Val(icmbAlignment.SelectedItem.tag)
                            End With
                        Case "tsDoubleClick"
                            With .Actions.OnDoubleClick
                                .TargetFrame = cmbTargetFrame.Text
                                .Type = cmbActionType.ListIndex
                                .url = txtURL.Text
                                .TargetMenu = cmbTargetMenu.ListIndex + 1
                                If .Type = atcCascade Then .TargetMenuAlignment = Val(icmbAlignment.SelectedItem.tag)
                            End With
                    End Select
                End With
            Else
                MenuCmds(sId).Compile = -chkCompile.Value
            End If
        End If
    End If
    
    SaveExpansions
    SaveState Split(StateInfo, cSep)(0) + ItmName + Split(StateInfo, cSep)(1)
    IsUpdating = False

End Sub

Private Sub UpdateEasyLink()

    Dim IsG As Boolean
    Dim IsC As Boolean
    Dim IsS As Boolean
    
    If tvMenus.SelectedItem Is Nothing Or IsTBMapSel Then
        Exit Sub
    Else
        IsG = IsGroup(tvMenus.SelectedItem.key)
        IsC = IsCommand(tvMenus.SelectedItem.key)
        IsS = IsSeparator(tvMenus.SelectedItem.key)
    End If
    
    If IsS Then Exit Sub
    
    Dim sUrl As String
    Dim ss As Integer
    Dim sl As Integer
    
    ss = txtEasyLink.SelStart
    sl = txtEasyLink.SelLength
    sUrl = txtEasyLink.Text
    
    tsCmdType.Tabs(2).Selected = True
    tsCmdType_Click
      
    If IsC Then
        With MenuCmds(GetID)
            If txtEasyLink.Text = "" Then
                cmbActionType.ListIndex = 0
            Else
                If .Actions.onclick.Type = atcNone Then cmbActionType.ListIndex = 1
            End If
        End With
    Else
        With MenuGrps(GetID)
            If txtEasyLink.Text = "" Then
                cmbActionType.ListIndex = 0
            Else
                If .Actions.onclick.Type = atcNone Then cmbActionType.ListIndex = 1
            End If
        End With
    End If
    
    txtURL.Text = sUrl
    If txtEasyLink.Text <> sUrl Then
        txtEasyLink.Text = sUrl
        txtEasyLink.SelStart = ss
        txtEasyLink.SelLength = sl
    End If

End Sub

Private Sub txtEasyLink_KeyUp(KeyCode As Integer, Shift As Integer)

    UpdateEasyLink

End Sub

Private Sub txtStatus_Change()

    If IsUpdating Then Exit Sub
    DontRefreshMap = True
    UpdateItemData GetLocalizedStr(234) + " " + cSep + " " + GetLocalizedStr(114)
    
End Sub

Private Sub txtURL_Change()

    If IsUpdating Then Exit Sub
    DontRefreshMap = True
    SetBookmarkState
    UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(294)
    
    Dim si As Node
    Dim n As Node
    Dim id As Integer
    
    If InMapMode Then
        Set si = tvMenus.SelectedItem
        Set n = tvMapView.SelectedItem
        id = GetID(si)
        If IsGroup(si.key) Then
            If MemberOf(id) = 0 Then
                n.ForeColor = GetItemColorByURL(MenuGrps(id).Actions, Preferences.GroupStyle.Color)
            Else
                n.ForeColor = GetItemColorByURL(MenuGrps(id).Actions, Preferences.ToolbarItemStyle.Color)
            End If
        Else
            n.ForeColor = GetItemColorByURL(MenuCmds(id).Actions, Preferences.CommandStyle.Color)
        End If
    End If
    
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)

    If Len(txtURL.Text) > 11 Then
        If Not (IsExternalLink(txtURL.Text) Or UsesProtocol(txtURL.Text)) Then
            With TipsSys
                .CanDisable = True
                .TipTitle = GetLocalizedStr(877)
                .Tip = GetLocalizedStr(878)
                .Show
            End With
        End If
    End If

End Sub

Private Sub udLeading_Change()

    MenuGrps(GetID).Leading = udLeading.Value
    
    shpSpc(1).Top = shpSpc(0).Top - shpSpc(0).Height - udLeading.Value * 15
    shpSpc(2).Top = shpSpc(0).Top + shpSpc(0).Height + udLeading.Value * 15
    
    If Not IsUpdating Then
        SaveState GetLocalizedStr(189) + " " + MenuGrps(GetID).Name + " " + GetLocalizedStr(295)
        UpdateLivePreview
    End If

End Sub

Private Sub MimicHoverButtonHover(State As Boolean)

    State = State And udLeading.Enabled

    ln1(0).Visible = State
    ln1(1).Visible = State
    ln1(2).Visible = State
    ln1(3).Visible = State
    
End Sub

Private Sub udLeading_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    MimicHoverButtonHover True

End Sub

Friend Sub SelectiveCopy(Optional ByVal ShowDlg As Boolean = True)

    On Error Resume Next

    If tvMenus.SelectedItem Is Nothing Then Exit Sub
    
    With frmSelectiveCopyPaste
        If IsTBMapSel Then
            dmbClipboard.TBContents = Project.Toolbars(ToolbarIndexByKey(frmMain.tvMapView.SelectedItem.key))
            dmbClipboard.ObjSrc = docToolbar
            .lblmsg.caption = "Select the styles to copy from the Toolbar" + " " + frmMain.tvMapView.SelectedItem.tag
        Else
            If IsGroup(tvMenus.SelectedItem.key) Then
                dmbClipboard.GrpContents = MenuGrps(GetID)
                dmbClipboard.ObjSrc = docGroup
                .lblmsg.caption = GetLocalizedStr(296) + " " + NiceGrpCaption(GetID)
            Else
                dmbClipboard.CmdContents = MenuCmds(GetID)
                If IsSeparator(tvMenus.SelectedItem.key) Then
                    dmbClipboard.ObjSrc = docSeparator
                    .lblmsg.caption = GetLocalizedStr(297) + " " + NiceGrpCaption(GetID(tvMenus.SelectedItem.parent))
                Else
                    dmbClipboard.ObjSrc = docCommand
                    .lblmsg.caption = GetLocalizedStr(298) + " " + NiceCmdCaption(GetID)
                End If
            End If
        End If
        
        .caption = GetLocalizedStr(299)
        .imgCopy.Visible = True
        .imgPaste.Visible = False
        .frmPasteOptions.Visible = False
        .uc3DLineSep.Visible = True
        .CreateNodes
        .Visible = True
        .Form_Resize
        .Visible = False
        If ShowDlg Then
            .Show vbModal
        End If
    End With
    
    If InMapMode Then
        SetCtrlFocus tvMapView
    Else
        SetCtrlFocus tvMenus
    End If

End Sub

Friend Sub SelectivePaste(Optional ByVal ForcePasteOption As Integer = -1)

    Dim ValidData As Boolean
    
    On Error Resume Next

    If tvMenus.SelectedItem Is Nothing Then Exit Sub
    
    ValidData = (IsGroup(tvMenus.SelectedItem.key) And dmbClipboard.ObjSrc = docGroup)
    ValidData = (IsCommand(tvMenus.SelectedItem.key) And dmbClipboard.ObjSrc = docCommand) Or ValidData
    ValidData = (IsSeparator(tvMenus.SelectedItem.key) And dmbClipboard.ObjSrc = docSeparator) Or ValidData
    ValidData = (IsTBMapSel And dmbClipboard.ObjSrc = docToolbar) Or ValidData

    If Not ValidData Then
        MsgBox GetLocalizedStr(603), vbInformation + vbOKOnly, GetLocalizedStr(604)
        Exit Sub
    End If
    
    With frmSelectiveCopyPaste
        .caption = GetLocalizedStr(300)
        .lblmsg.caption = GetLocalizedStr(301)
        
        If IsTBMapSel Then
            .txtOp(0).Text = "Paste to the toolbar " + tvMapView.SelectedItem.tag
            .txtOp(1).Text = ""
            .opPaste(1).Enabled = False
            If .opPaste(1).Value Then .opPaste(0).Value = True
            .txtOp(2).Text = "Paste to all toolbars"
            .opPaste(3).Enabled = True
            .cmdAdvanced.Enabled = True
            .cmdAdvanced.Visible = False
        Else
            If IsGroup(tvMenus.SelectedItem.key) Then
                .txtOp(0).Text = GetLocalizedStr(302) + " " + NiceGrpCaption(GetID)
                .txtOp(1).Text = ""
                .opPaste(1).Enabled = False
                If .opPaste(1).Value Then .opPaste(0).Value = True
                .txtOp(2).Text = GetLocalizedStr(303)
                .opPaste(3).Enabled = True
                .cmdAdvanced.Enabled = True
                .cmdAdvanced.Visible = True
            Else
                .opPaste(1).Enabled = True
                If IsSeparator(tvMenus.SelectedItem.key) Then
                    .txtOp(0).Text = GetLocalizedStr(304)
                    .txtOp(1).Text = GetLocalizedStr(305) + " " + NiceGrpCaption(GetID(tvMenus.SelectedItem.parent)) + " " + GetLocalizedStr(306)
                    .txtOp(2).Text = GetLocalizedStr(307)
                    .opPaste(3).Enabled = False
                    .cmdAdvanced.Enabled = False
                Else
                    .txtOp(0).Text = GetLocalizedStr(308) + " " + NiceCmdCaption(GetID)
                    .txtOp(1).Text = GetLocalizedStr(309) + " " + NiceGrpCaption(GetID(tvMenus.SelectedItem.parent)) + " " + GetLocalizedStr(306)
                    .txtOp(2).Text = GetLocalizedStr(310)
                    .opPaste(3).Enabled = True
                    .cmdAdvanced.Enabled = True
                End If
            End If
        End If
        .imgCopy.Visible = False
        .imgPaste.Visible = True
        .frmPasteOptions.Visible = True
        .uc3DLineSep.Visible = False
        .RemoveUnselectedNodes
        If ForcePasteOption <> -1 Then
            frmSelectiveCopyPaste.opPaste(ForcePasteOption).Value = True
        End If
        If .tvProperties.Nodes.Count = 0 Then
            MsgBox GetLocalizedStr(311), vbInformation + vbOKOnly, GetLocalizedStr(604)
        Else
            .Visible = True
            .Form_Resize
            .Visible = False
            .Show vbModal
        End If
    End With
    
    UpdateControls
    RefreshMap
    
    If InMapMode Then
        SetCtrlFocus tvMapView
    Else
        SetCtrlFocus tvMenus
    End If

End Sub

Private Sub SetDefBrowserIcon()

    Dim bc As String
    Dim bci As String
    Dim lcText As String
    
    On Error Resume Next
    
    lcText = GetLocalizedStr(657) + " "

    bci = Val(GetSetting(App.EXEName, "Browsers", "Default", 1)) - 1
    If bci = 0 Then
        tbMenu2.Buttons("tbPreview").Image = ilIcons.ListImages("mnuToolsPreview").Index
        tbMenu2.Buttons("tbPreview").ToolTipText = lcText + "Internet Explorer (" + GetLocalizedStr(658) + ")"
    Else
        bc = GetSetting(App.EXEName, "Browsers", "Command" & bci)
        If FileExists(bc) Then
            GetIcon picIcon, bc
            Set picIcon.Picture = picIcon.Image
            
            If ilIcons.ListImages(bc) Is Nothing Then
                Err.Clear
                vlmCtrl.SetImageList Nothing
                ilIcons.ListImages.Add , bc, picIcon.Picture
                vlmCtrl.SetImageList ilIcons
            End If
            tbMenu2.Buttons("tbPreview").Image = ilIcons.ListImages(bc).Index
        Else
            tbMenu2.Buttons("tbPreview").Image = ilIcons.ListImages("EmptyIcon").Index
        End If
        tbMenu2.Buttons("tbPreview").ToolTipText = lcText + GetSetting(App.EXEName, "Browsers", "Name" & bci)
    End If
    
    If Err.number <> 0 Then tbMenu2.Buttons("tbPreview").Image = ilIcons.ListImages("mnuToolsPreview").Index
  
    AddBrowsers2Menu
    
End Sub

Private Sub AddBrowsers2Menu()

    Dim i As Integer
    Dim DefBrowser As Integer
    Dim bFile As String
    
    DefBrowser = Val(GetSetting(App.EXEName, "Browsers", "Default", 1)) - 1
    
    While mnuBrowsersList.Count > 1
        Unload mnuBrowsersList(mnuBrowsersList.Count - 1)
    Wend
    
    i = 1
    Do Until LenB(GetSetting(App.EXEName, "Browsers", "Name" & i, "")) = 0
        bFile = GetSetting(App.EXEName, "Browsers", "Command" & i)
        
        Load mnuBrowsersList(i)
        mnuBrowsersList(i).Enabled = FileExists(bFile)
        mnuBrowsersList(i).caption = GetSetting(App.EXEName, "Browsers", "Name" & i) + IIf(mnuBrowsersList(i).Enabled, " " + GetFileVersion(bFile, True), "")
        mnuBrowsersList(i).Checked = (DefBrowser = i)
        mnuBrowsersList(i).tag = bFile
        i = i + 1
    Loop
    
    If Not IsDebug Then xMenu.Initialize Me

End Sub

Private Sub SaveExpansions()

    Dim nNode As Node

    Project.NodeExpStatus = ""
    If tvMapView.Nodes.Count < 2000 Then
        For Each nNode In tvMapView.Nodes
            With nNode
                If .children > 0 And .Expanded Then
                    Project.NodeExpStatus = Project.NodeExpStatus + "+" + .FullPath
                End If
            End With
        Next nNode
    End If
    Project.NodeExpStatus = Project.NodeExpStatus + "+"
    
End Sub

Friend Sub RefreshMap(Optional UseLastExp As Boolean)

    Dim g As Integer
    Dim t As Integer
    Dim nNode As Node
    Dim mt As Integer
    Dim tt As Integer
    
    Dim v() As Boolean
    
    If Not InMapMode Then Exit Sub
    If IsReplacing Then Exit Sub
    
    If DontRefreshMap Then
        DontRefreshMap = False
        If Not tvMapView.SelectedItem Is Nothing And Not IsTBMapSel Then
            If LenB(txtCaption.Text) <> 0 Then
                tvMapView.SelectedItem.Text = txtCaption.Text
                Exit Sub
            End If
        End If
    End If
    
    IsRefreshingMap = True
    
    ReDim v(tvMapView.Nodes.Count)
    For Each nNode In tvMapView.Nodes
        v(nNode.Index) = nNode.Visible
    Next nNode
    
    If Not UseLastExp Then
        If KeepExpansions Then SaveExpansions
    End If
    
    tvMapView.Nodes.Clear
    
    tt = UBound(Project.Toolbars)
    If CreateToolbar Then
        For t = 1 To tt
            If UBound(Project.Toolbars(t).Groups) = 0 Then
                AddTBNode Project.Toolbars(t).Name
            Else
                For g = 1 To UBound(Project.Toolbars(t).Groups)
                    RenderGroup GetIDByName(Project.Toolbars(t).Groups(g)) ', , t
                Next g
            End If
        Next t
        For g = 1 To UBound(MenuGrps)
            mt = MemberOf(g)
            If (mt = 0 Or mt > tt) And Not IsSubMenu(g) Then RenderGroup g
        Next g
        AddTBNode
    Else
        For g = 1 To UBound(MenuGrps)
            RenderGroup g
        Next g
    End If
    
    Set nNode = tvMapView.SelectedItem
    If nNode Is Nothing And Not tvMenus.SelectedItem Is Nothing Then
        For Each nNode In tvMapView.Nodes
            If LenB(LastSelNode) = 0 Then
                If nNode.tag = tvMenus.SelectedItem.key Then
                    nNode.Selected = True
                    nNode.EnsureVisible
                    Exit For
                End If
            Else
                If nNode.FullPath = LastSelNode Then
                    nNode.Selected = True
                    nNode.EnsureVisible
                    Exit For
                End If
            End If
        Next nNode
    End If
    
    If tvMapView.SelectedItem Is Nothing Then SynchViews
    
    On Error Resume Next
    For g = UBound(v) To 1 Step -1
        If v(g) Then tvMapView.Nodes(g).EnsureVisible
    Next g
    
    KeepExpansions = True
    
    Me.Enabled = True
    Me.MousePointer = vbDefault
    
    IsRefreshingMap = False

End Sub

Private Function ForceExpanded(nNode As Node) As Boolean

    ForceExpanded = (InStr(Project.NodeExpStatus, "+" + nNode.FullPath + "+") > 0)

End Function

Private Function AddTBNode(Optional TBName As String) As Node

    Dim nNode As Node
    
    On Error Resume Next
    
    If LenB(TBName) = 0 Then TBName = GetLocalizedStr(795)

    Set nNode = tvMapView.Nodes("TBK" + TBName)
    If nNode Is Nothing Then
        Set nNode = tvMapView.Nodes.Add(, , "TBK" + TBName, TBName, IconIndex("Toolbar Item"))
        nNode.tag = TBName
        nNode.Expanded = ForceExpanded(nNode)
        
        nNode.ForeColor = Preferences.ToolbarStyle.Color
        nNode.Bold = Preferences.ToolbarStyle.Font.FontBold
        
        With Project.Toolbars(ToolbarIndexByKey(nNode.key))
            If .IsTemplate Then nNode.ForeColor = vbBlue
            If Not .Compile Then nNode.ForeColor = Preferences.NoCompileItem
        End With
    End If
    
    If TBName = GetLocalizedStr(795) Then nNode.ForeColor = &H808080
    
    Set AddTBNode = nNode

End Function

Private Sub RenderGroup(g As Integer, Optional rc As Node, Optional t As Integer = 0, Optional ExpandGroup As Boolean = False)

    Dim c As Integer
    Dim ng As Node
    Dim nc As Node
    Static rec As Integer
    Dim tbn As Node
    Dim rg As Integer
    Dim ForceFullRender As Boolean
    Dim IsT As Boolean
    Dim moidx As Integer
    
    On Error Resume Next
    
    rec = rec + 1

    With MenuGrps(g)
        moidx = MemberOf(g)
        IsT = .IsTemplate
        If rc Is Nothing Then
            If CreateToolbar And moidx > 0 Then
                If t = 0 Then t = moidx
                Set tbn = AddTBNode(Project.Toolbars(t).Name)
            Else
                Set tbn = AddTBNode
            End If
            Set ng = tvMapView.Nodes.Add(tbn, tvwChild, , IIf(LenB(.caption) = 0, "[" + .Name + "]", xUNI2Unicode(.caption)))
            ng.Expanded = True
        Else
            Set ng = tvMapView.Nodes.Add(rc, tvwChild, , IIf(LenB(.caption) = 0, "[" + .Name + "]", xUNI2Unicode(.caption)), tvMenus.Nodes("G" & g).Image)
        End If
        ng.tag = "G" & g
        ng.Image = tvMenus.Nodes(ng.tag).Image
        
        If Not .Compile Or ng.parent.ForeColor = Preferences.NoCompileItem Then
            ng.ForeColor = Preferences.NoCompileItem
        Else
            If t > 0 Then
                ng.ForeColor = IIf(.disabled, Preferences.DisabledItem, GetItemColorByURL(.Actions, Preferences.ToolbarItemStyle.Color))
                ng.Bold = Preferences.ToolbarItemStyle.Font.FontBold
            Else
                ng.ForeColor = IIf(.disabled, Preferences.DisabledItem, GetItemColorByURL(.Actions, Preferences.GroupStyle.Color))
                ng.Bold = Preferences.GroupStyle.Font.FontBold
            End If
        End If
        
        ng.Expanded = ExpandGroup Or ForceExpanded(ng)
        If .IsTemplate Then ng.ForeColor = vbBlue
        If rec > 20 Then
            Set nc = tvMapView.Nodes.Add(ng, tvwChild, , "(too many recursions)")
            nc.ForeColor = vbRed
        Else
            ForceFullRender = (ng.parent.parent Is Nothing)
            rg = GetRealSubGroup(g, True)
            For c = 1 To UBound(MenuCmds)
                With MenuCmds(c)
                    If .parent = rg Then
                        'Set nc = tvMapView.Nodes.Add(ng, tvwChild, , IIf(.Name = "[SEP]", String$(10, "-"), IIf(lenb(.Caption )=0, "[" + .Name + "]", .Caption)), IIf(.Name = "[SEP]", IconIndex("Separator"), GenCmdIcon(c)))
                        Set nc = tvMapView.Nodes.Add(ng, tvwChild, , IIf(.Name = "[SEP]", String$(10, "-"), IIf(LenB(.caption) = 0, "[" + .Name + "]", xUNI2Unicode(.caption))))
                        nc.tag = IIf(.Name = "[SEP]", "S", "C") & c
                        nc.Image = tvMenus.Nodes(nc.tag).Image
                        
                        If IsT Then
                            nc.ForeColor = ng.ForeColor
                        Else
                            If Not .Compile Or nc.parent.ForeColor = Preferences.NoCompileItem Then
                                nc.ForeColor = Preferences.NoCompileItem
                            Else
                                nc.ForeColor = IIf(.disabled, Preferences.DisabledItem, GetItemColorByURL(.Actions, Preferences.CommandStyle.Color))
                            End If
                        End If
                        nc.Bold = Preferences.CommandStyle.Font.FontBold
                        
                        nc.Expanded = ForceExpanded(nc)
                        If .Actions.onmouseover.Type = atcCascade Then RenderGroup .Actions.onmouseover.TargetMenu, nc
                        If .Actions.onclick.Type = atcCascade Then RenderGroup .Actions.onclick.TargetMenu, nc
                        If .Actions.OnDoubleClick.Type = atcCascade Then RenderGroup .Actions.OnDoubleClick.TargetMenu, nc
                        If Not ng.Expanded And Not ForceFullRender Then Exit For
                    End If
                End With
            Next c
        End If
        If NagScreenIsVisible Then DoEvents
    End With
    
    rec = rec - 1

End Sub

Private Function GetItemColorByURL(a As ActionEvents, DefaultColor As OLE_COLOR) As OLE_COLOR

    GetItemColorByURL = IIf(IsLinkValid(a), DefaultColor, Preferences.BrokenLink)
    
End Function

Friend Function IsLinkValid(a As ActionEvents) As Boolean

    Dim p1 As Long
    Dim p2 As Long
    Dim c As ConfigDef
    Dim url As String
    Dim rw As String
    
    rw = GetRealLocal.RootWeb
    url = a.onclick.url
    
    If LenB(url) = 0 Then
        IsLinkValid = True
    Else
        If FolderExists(rw) Then
            If a.onclick.Type = atcURL Or a.onclick.Type = atcNewWindow Then
                If UsesProtocol(url) Then
                    IsLinkValid = True
                Else
                    c = Project.UserConfigs(Project.DefaultConfig)
                    If c.Type = ctcRemote Then
                        If LCase(Left(url, Len(c.RootWeb))) <> LCase(c.RootWeb) Then
                            IsLinkValid = True
                            Exit Function
                        End If
                        If Left(rw, 2) = "\\" Then
                            url = SetSlashDir(rw + FixURL(url), sdBack)
                            url = "\" + Replace(url, "\\", "\")
                        Else
                            url = Replace(SetSlashDir(rw + FixURL(url), sdBack), "\\", "\")
                        End If
                    ElseIf IsExternalLink(url) Then
                        IsLinkValid = True
                        Exit Function
                    End If
                    p1 = InStr(url, "#")
                    p2 = InStr(url, "?")
                    If p1 = 0 And p2 <> 0 Then p1 = p2
                    If p1 > 0 Then url = Left(url, p1 - 1)
                    IsLinkValid = FileExists(url)
                End If
            Else
                IsLinkValid = True
            End If
        Else
            IsLinkValid = True
        End If
    End If

End Function

Private Sub LocalizeUI()

    lblDataTitle.caption = GetLocalizedStr(102)
    lblCaption.caption = GetLocalizedStr(103)
    chkEnabled.caption = GetLocalizedStr(104)
    chkCompile.caption = Replace(GetLocalizedStr(161), "&", "")
    
    tsCmdType.Tabs(1).caption = GetLocalizedStr(105)
    tsCmdType.Tabs(2).caption = GetLocalizedStr(106)
    tsCmdType.Tabs(3).caption = GetLocalizedStr(107)
    
    lblStatus.caption = GetLocalizedStr(114)
    lblAlignment.caption = GetLocalizedStr(115)
    lblActionType.caption = GetLocalizedStr(108)
    lblLayout.caption = GetLocalizedStr(668)
    lblASVertical.caption = GetLocalizedStr(210)
    lblASHorizontal.caption = GetLocalizedStr(211)
    
    lblTSViewsMap.caption = GetLocalizedStr(758)
    lblTSViewsNormal.caption = GetLocalizedStr(759)

    'tbMenu Toolbar
    tbMenu.Buttons("tbNew").ToolTipText = GetLocalizedStr(127)
    tbMenu.Buttons("tbOpen").ToolTipText = GetLocalizedStr(128)
    tbMenu.Buttons("tbSave").ToolTipText = GetLocalizedStr(130)
    tbMenu.Buttons("tbAddGrp").ToolTipText = GetLocalizedStr(147)
    tbMenu.Buttons("tbAddSubGrp").ToolTipText = GetLocalizedStr(760)
    tbMenu.Buttons("tbAddCmd").ToolTipText = GetLocalizedStr(148)
    tbMenu.Buttons("tbAddSep").ToolTipText = GetLocalizedStr(149)
    tbMenu.Buttons("tbCopy").ToolTipText = GetLocalizedStr(138)
    tbMenu.Buttons("tbPaste").ToolTipText = GetLocalizedStr(139)
    tbMenu.Buttons("tbFind").ToolTipText = GetLocalizedStr(140)
    tbMenu.Buttons("tbRemove").ToolTipText = GetLocalizedStr(143)
    
    'tbCmd Toolbar
    tbCmd.Buttons("tbColor").ToolTipText = GetLocalizedStr(150)
    tbCmd.Buttons("tbFont").ToolTipText = GetLocalizedStr(151)
    tbCmd.Buttons("tbCursor").ToolTipText = GetLocalizedStr(152)
    tbCmd.Buttons("tbImage").ToolTipText = GetLocalizedStr(153)
    picLeading.ToolTipText = GetLocalizedStr(295)
    tbCmd.Buttons("tbMargins").ToolTipText = GetLocalizedStr(154)
    tbCmd.Buttons("tbFX").ToolTipText = GetLocalizedStr(155)
    tbCmd.Buttons("tbSelFX").ToolTipText = GetLocalizedStr(984)
    tbCmd.Buttons("tbSound").ToolTipText = GetLocalizedStr(156)
    tbCmd.Buttons("tbUp").ToolTipText = GetLocalizedStr(655)
    tbCmd.Buttons("tbDown").ToolTipText = GetLocalizedStr(656)
    
    'tbMenu2 Toolbar
    tbMenu2.Buttons("tbCompile").ToolTipText = GetLocalizedStr(161)
    tbMenu2.Buttons("tbPublish").ToolTipText = GetLocalizedStr(163)
    tbMenu2.Buttons("tbCompile").ToolTipText = GetLocalizedStr(161)
    tbMenu2.Buttons("tbProperties").ToolTipText = GetLocalizedStr(133)
    tbMenu2.Buttons("tbHotSpotEditor").ToolTipText = GetLocalizedStr(166)
    'mnuBrowsersSetDefBrowser.Caption = GetLocalizedStr(160)
    
    LocalizeMenuItems
    
    frmFind.LocalizeUI
    frmSelectiveCopyPaste.LocalizeUI

End Sub

Private Sub LocalizeMenuItems()

    mnuFile.caption = GetLocalizedStr(126)
        mnuFileNew.caption = GetLocalizedStr(127)
            mnuFileNewEmpty.caption = GetLocalizedStr(804)
            mnuFileNewFromPreset.caption = GetLocalizedStr(805)
            mnuFileNewFromWizard.caption = GetLocalizedStr(806)
            mnuFileNewFromDir.caption = GetLocalizedStr(833) + "..."
        mnuFileOpen.caption = GetLocalizedStr(128)
        mnuFileOpenRecent.caption = GetLocalizedStr(129)
        mnuFileOpenRecentOP.caption = GetLocalizedStr(615)
        mnuFileSave.caption = GetLocalizedStr(130)
        mnuFileSaveAs.caption = GetLocalizedStr(131)
        mnuFileSaveAsPreset.caption = GetLocalizedStr(834) + "..."
        mnuFileSubmitPreset.caption = GetLocalizedStr(913) + "..."
        mnuFileExportHTML.caption = GetLocalizedStr(535)
        mnuFileProjProp.caption = GetLocalizedStr(133)
        mnuFileExit.caption = GetLocalizedStr(134)
        
    mnuEdit.caption = GetLocalizedStr(135)
        mnuEditUndo.caption = GetLocalizedStr(136)
        mnuEditRedo.caption = GetLocalizedStr(137)
        mnuEditCopy.caption = GetLocalizedStr(138)
        mnuEditPaste.caption = GetLocalizedStr(139)
        mnuEditFind.caption = GetLocalizedStr(140)
        mnuEditFindNext.caption = GetLocalizedStr(141)
        mnuEditFindReplace.caption = GetLocalizedStr(142)
        mnuEditDelete.caption = GetLocalizedStr(143)
        mnuEditRename.caption = GetLocalizedStr(144)
        mnuEditPreferences.caption = GetLocalizedStr(145)
        
    mnuMenu.caption = GetLocalizedStr(146)
        mnuMenuAddToolbar.caption = GetLocalizedStr(809)
        mnuMenuAddGroup.caption = GetLocalizedStr(147)
        mnuMenuAddSubGroup.caption = GetLocalizedStr(760)
        mnuMenuAddCommand.caption = GetLocalizedStr(148)
        mnuMenuAddSeparator.caption = GetLocalizedStr(149)
        mnuMenuColor.caption = GetLocalizedStr(150)
        mnuMenuFont.caption = GetLocalizedStr(151)
        mnuMenuCursor.caption = GetLocalizedStr(152)
        mnuMenuImage.caption = GetLocalizedStr(153)
        mnuMenuMargins.caption = GetLocalizedStr(154)
        mnuMenuSelFX.caption = GetLocalizedStr(984) + "..."
        mnuMenuSFX.caption = GetLocalizedStr(155)
        mnuMenuLength.caption = GetLocalizedStr(811)
        mnuMenuToolbarProperties.caption = GetLocalizedStr(816)
        mnuMenuCopy.caption = GetLocalizedStr(138)
        mnuMenuPaste.caption = GetLocalizedStr(139)
        mnuMenuDelete.caption = GetLocalizedStr(143)
        mnuMenuRename.caption = GetLocalizedStr(144)
        
    mnuTools.caption = GetLocalizedStr(157)
        mnuToolsPreview.caption = GetLocalizedStr(158)
        mnuToolsLivePreview.caption = GetLocalizedStr(159)
        mnuToolsSetDefaultBrowser.caption = GetLocalizedStr(160)
        mnuToolsCompile.caption = GetLocalizedStr(161)
        mnuToolsPublish.caption = GetLocalizedStr(163)
        mnuToolsInstallMenusA.caption = GetLocalizedStr(761)
            mnuToolsInstallMenusAILC.caption = GetLocalizedStr(164)
            mnuToolsInstallMenusAIFLC.caption = GetLocalizedStr(165)
            mnuToolsInstallMenusAIRLC.caption = GetLocalizedStr(762) + "..."
        mnuToolsInstallMenus.caption = GetLocalizedStr(761) + "..."
        mnuToolsDefaultConfig.caption = GetLocalizedStr(282) + "..."
        mnuToolsHotSpotsEditor.caption = GetLocalizedStr(166)
        mnuToolsToolbarsEditor.caption = GetLocalizedStr(167)
        #If DEVVER = 1 Then
            mnuToolsDynAPI.caption = "DynAPI"
        #End If
        mnuToolsAddInEditor.caption = GetLocalizedStr(168)
        mnuToolsApplyStyle.caption = GetLocalizedStr(803)
        mnuToolsSecProj.caption = GetLocalizedStr(825) + "..."
        mnuToolsExtractIcon.caption = GetLocalizedStr(835) + "..."
        mnuToolsReports.caption = GetLocalizedStr(937)
        mnuToolsReport.caption = GetLocalizedStr(169)
        mnuToolsBrokenLinks.caption = GetLocalizedStr(925) + "..."
    
    mnuHelp.caption = GetLocalizedStr(170)
        mnuHelpContents.caption = GetLocalizedStr(171)
        mnuHelpTutorials.caption = GetLocalizedStr(172)
        mnuHelpSearch.caption = GetLocalizedStr(693)
        mnuHelpUpgrade.caption = GetLocalizedStr(173)
        mnuHelpXFX.caption = GetLocalizedStr(174)
            mnuHelpXFXHomePage.caption = GetLocalizedStr(312)
            mnuHelpXFXSupport.caption = GetLocalizedStr(313)
            mnuHelpXFXFAQ.caption = GetLocalizedStr(314)
            mnuHelpXFXPublicForum.caption = GetLocalizedStr(315)
            mnuHelpXFXNews.caption = GetLocalizedStr(316)
        mnuHelpAbout.caption = GetLocalizedStr(175)
        
    mnuBrowsersList(0).caption = GetLocalizedStr(817)
    
    mnuPPShortcutsGeneral.caption = GetLocalizedStr(321) + "..."
    mnuPPShortcutsConfig.caption = GetLocalizedStr(322) + "..."
    mnuPPShortcutsGlobal.caption = GetLocalizedStr(727) + "..."
    mnuPPShortcutsAdvanced.caption = GetLocalizedStr(325) + "..."
    
    mnuTBEShortcutsGeneral.caption = GetLocalizedStr(321) + "..."
    mnuTBEShortcutsAppearance.caption = GetLocalizedStr(352) + "..."
    mnuTBEShortcutsPositioning.caption = GetLocalizedStr(353) + "..."
    mnuTBEShortcutsEffects.caption = GetLocalizedStr(832) + "..."
    mnuTBEShortcutsAdvanced.caption = GetLocalizedStr(325) + "..."
    
    If IsDEMO Then
        mnuRegister.caption = GetLocalizedStr(176)
        mnuRegisterUnlock.caption = GetLocalizedStr(177)
        mnuRegisterBuy.caption = GetLocalizedStr(178)
    End If
    
    If Not IsDebug Then xMenu.Initialize Me
        
End Sub

Private Sub vlmCtrl_DrawMenuItem(ByVal aMenuItem As VLMnuPlus.CMenuItem, ByVal menuDC As Long, ByVal bSelected As Boolean, ByVal cx As Long, ByVal cy As Long, bUserDrawn As Boolean)

'    If xMenu.Name(aMenuItem.Caption) = "mnuMenu" Then
'        SetupMenuMenu Nothing
'    End If
    
End Sub

Private Sub vlmCtrl_SetMenuItemAttributes(ByVal aMenuItem As VLMnuPlus.CMenuItem)

    Dim mName As String
    Dim i As Integer
    Dim DefBrowser As Integer

    On Local Error Resume Next
    
    If aMenuItem.IsTopLevel Then Exit Sub
    If aMenuItem.IsSeparator Then Exit Sub

    mName = xMenu.Name(aMenuItem.caption)
    If LenB(mName) = 0 Then Exit Sub
    
    If Left(mName, 18) = "mnuFileOpenRecentR" Then
        Set aMenuItem.Picture = ilIcons.ListImages("mnuFileRF").ExtractIcon
        Exit Sub
    End If
    
    If Left(mName, 15) = "mnuBrowsersList" Then
        DefBrowser = GetSetting(App.EXEName, "Browsers", "Default", 1) - 1
        i = Split(mName, "|")(1)
        If i = 0 Then
            Set aMenuItem.Picture = ilIcons.ListImages("mnuToolsPreview").ExtractIcon
        Else
            GetIcon picIcon, mnuBrowsersList(i).tag
            Set picIcon.Picture = picIcon.Image
            Set aMenuItem.Picture = picIcon.Picture
        End If
        aMenuItem.CaptionFont.Bold = (i = DefBrowser)
        Exit Sub
    End If

    i = xMenu.IconIndex(mName)
    If i = 0 Then
        i = GetIconIdx(mName)
        xMenu.IconIndex(mName) = i
    End If
    Set aMenuItem.Picture = ilIcons.ListImages(i).ExtractIcon
    
End Sub

Private Function GetIconIdx(mName As String) As Integer

    Dim sk As String
    Dim s() As String
    Dim lImg As ListImage
    Dim i As Integer

    For Each lImg In ilIcons.ListImages
        sk = lImg.key + "|"
        If InStr(1, sk, mName, vbTextCompare) Then
            s = Split(sk, "|")
            For i = 0 To UBound(s) - 1
                If s(i) = mName Then
                    GetIconIdx = lImg.Index
                    Exit Function
                End If
            Next i
        End If
    Next lImg

End Function

Private Sub vlmCtrl_SetMenuItemSize(ByVal aMenuItem As VLMnuPlus.CMenuItem, ItemWidth As Long, ItemHeight As Long)

    If aMenuItem.IsSeparator Then ItemHeight = 5

End Sub

Private Sub SetTBIcons()

    With tbMenu
        .ImageList = ilIcons
        
        .Buttons("tbNew").Image = GetIconIdx("mnuFileNew")
        .Buttons("tbOpen").Image = GetIconIdx("mnuFileOpen")
        .Buttons("tbSave").Image = GetIconIdx("mnuFileSave")
        
        .Buttons("tbAddGrp").Image = GetIconIdx("mnuMenuAddGroup")
        .Buttons("tbAddCmd").Image = GetIconIdx("mnuMenuAddCommand")
        .Buttons("tbAddSubGrp").Image = GetIconIdx("mnuMenuAddSubGroup")
        .Buttons("tbAddSep").Image = GetIconIdx("mnuMenuAddSeparator")
        
        .Buttons("tbCopy").Image = GetIconIdx("mnuEditCopy")
        .Buttons("tbPaste").Image = GetIconIdx("mnuEditPaste")
        
        .Buttons("tbFind").Image = GetIconIdx("mnuEditFind")
        
        .Buttons("tbUndo").Image = GetIconIdx("mnuEditUndo")
        .Buttons("tbRedo").Image = GetIconIdx("mnuEditRedo")
        
        .Buttons("tbRemove").Image = GetIconIdx("mnuEditDelete")
    End With
    
    With tbCmd
        .ImageList = ilIcons
        
        .Buttons("tbColor").Image = GetIconIdx("mnuMenuColor")
        .Buttons("tbFont").Image = GetIconIdx("mnuMenuFont")
        .Buttons("tbCursor").Image = GetIconIdx("mnuMenuCursor")
        .Buttons("tbImage").Image = GetIconIdx("mnuMenuImage")
        .Buttons("tbImage").Image = GetIconIdx("mnuMenuImage")
        .Buttons("tbMargins").Image = GetIconIdx("mnuMenuMargins")
        .Buttons("tbFX").Image = GetIconIdx("mnuMenuSFX")
        .Buttons("tbSelFX").Image = GetIconIdx("mnuMenuSelFX")
        
        .Buttons("tbUp").Image = GetIconIdx("btnUp")
        .Buttons("tbDown").Image = GetIconIdx("btnDown")
    End With
    
    With tbMenu2
        .ImageList = ilIcons
        
        .Buttons("tbCompile").Image = GetIconIdx("mnuToolsCompile")
        .Buttons("tbPublish").Image = GetIconIdx("mnuToolsPublish")
        .Buttons("tbHotSpotEditor").Image = GetIconIdx("mnuToolsHotSpotsEditor")
        .Buttons("tbTBEditor").Image = GetIconIdx("mnuToolsToolbarsEditor")
        .Buttons("tbProperties").Image = GetIconIdx("mnuFileProjProp")
    End With

End Sub

Private Sub GetSystemCharset()

    Const DEFAULT_CHARSET = 1
    Const SYMBOL_CHARSET = 2
    Const SHIFTJIS_CHARSET = 128
    Const HANGEUL_CHARSET = 129
    Const CHINESEBIG5_CHARSET = 136
    Const CHINESESIMPLIFIED_CHARSET = 134
    
    FontCharSet = Val(GetSetting(App.EXEName, "Preferences", "FontCharSet", -1))
    If FontCharSet = -1 Then
        Select Case GetUserDefaultLCID
            Case &H404  ' Traditional Chinese
                FontCharSet = CHINESEBIG5_CHARSET
            Case &H411  ' Japan
                FontCharSet = SHIFTJIS_CHARSET
            Case &H412  'Korea UserLCID
                FontCharSet = HANGEUL_CHARSET
            Case &H804  ' Simplified Chinese
                FontCharSet = CHINESESIMPLIFIED_CHARSET
            Case Else   ' The other countries
                Err.Clear
                On Error Resume Next
                FontCharSet = Printer.Font.Charset
                If Err.number <> 0 Then FontCharSet = Me.Font.Charset
                Err.Clear
                On Error GoTo 0
        End Select
    End If
    
End Sub

Friend Sub DoDelayedLivePreview()

    tmrDelayedLivePreview.Enabled = False
    
    If UBound(MenuGrps) = 0 Or SelectiveLauncher = slcShowContextMenu Then Exit Sub

    tmrDelayedLivePreview.Enabled = True

End Sub

Friend Sub DoLivePreview(Optional altWB As WebBrowser, Optional DontRenderTBI As Boolean, Optional DontRenderGroup As Boolean, Optional ForceSelTB As Integer = 0)

    Dim g() As MenuGrp
    Dim c() As MenuCmd
    Dim i As Integer
    Dim p As ProjectDef
    Dim sStr As String
    Dim bgColor As Long
    Dim idx As Integer
    Dim IsG As Boolean
    Dim IsTB As Boolean
    Dim tbi As Integer
    Dim IsTBSel As Boolean
    Static IsBusy As Boolean
    Dim mpx As Integer
    Dim mpy As Integer
    
    Static lastWB As WebBrowser
    Static lastDontRenderTBI As Boolean
    Static lastForceSelTB As Integer
    
    Static lastHTML As String
    
    On Error GoTo ExitSub
    
    If Preferences.UseLivePreview = False Or UBound(MenuGrps) = 0 Or IsRenaming Or SelectiveLauncher = slcShowContextMenu Then Exit Sub
    
    If LivePreviewIsBusy Then
        If lastWB Is Nothing Then
            Set lastWB = altWB
            lastDontRenderTBI = DontRenderTBI
            lastForceSelTB = ForceSelTB
        End If
        
        DoDelayedLivePreview
        Exit Sub
    Else
        If Not lastWB Is Nothing Then
            Set altWB = lastWB
            DontRenderTBI = lastDontRenderTBI
            ForceSelTB = lastForceSelTB
        End If
    End If
    
    LivePreviewIsBusy = True
    
    IsTBSel = IsTBMapSel Or ForceSelTB <> 0
    If IsTBSel Then
        If ForceSelTB > 0 Then
            idx = ForceSelTB
        Else
            idx = ToolbarIndexByKey(tvMapView.SelectedItem.key)
        End If
        IsG = False
        IsTB = True
    Else
        If Not tvMenus.SelectedItem Is Nothing Then
            idx = GetID
            IsG = IsGroup(tvMenus.SelectedItem.key)
            IsTB = False
        End If
    End If
    
    If LivePreviewCharset = "" Then
        For i = 1 To UBound(cs)
            If cs(i).CodePage = Preferences.CodePage Then
                LivePreviewCharset = cs(i).WebCharset
                Exit For
            End If
        Next i
    End If
    
    p = Project
    With p
        .CompilehRefFile = False
        .CompileNSCode = False
        .DoFormsTweak = False
        .DWSupport = False
        .ImageReadySupport = False
        .KeyboardSupport = False
        .LotusDominoSupport = False
        .OPHelperFunctions = False
        .CustomOffsets = ""
        .AutoSelFunction = False
        .AddIn.Name = ""
        .FX = 0
        ReDim .SecondaryProjects(0)
    End With
    
    If IsTB Then
        g = MenuGrps
        For i = 1 To UBound(MenuGrps)
            g(i).Actions.onmouseover.Type = atcNone
            g(i).Actions.onclick.Type = atcNone
        Next i
        ReDim c(0)
        ReDim p.Toolbars(1)
        p.Toolbars(1) = Project.Toolbars(idx)
    Else
        ReDim g(1)
        If IsG Then
            ReDim c(0)
            g(1) = MenuGrps(idx)
            g(1).Name = "grpp1"
            g(1).IsContext = False
            
            If altWB Is Nothing Then
                DontRenderTBI = (sbLP_GroupMode.tag = "GRP")
            End If
            
            If DontRenderTBI Then
                tbi = 0
            Else
                tbi = MemberOf(idx)
            End If
            If tbi = 0 Then
                ReDim p.Toolbars(0)
                For i = 1 To UBound(MenuCmds)
                    If MenuCmds(i).parent = idx Then
                        ReDim Preserve c(UBound(c) + 1)
                        c(UBound(c)) = MenuCmds(i)
                        c(UBound(c)).parent = 1
                    End If
                Next i
            Else
                With p.MenusOffset
                    .RootMenusX = 0
                    .RootMenusY = 0
                    .SubMenusX = 0
                    .SubMenusY = 0
                End With
                
                ReDim p.Toolbars(1)
                p.Toolbars(1) = Project.Toolbars(tbi)
                ReDim p.Toolbars(1).Groups(1)
                p.Toolbars(1).Groups(1) = "grpp1"
                p.Toolbars(1).Condition = "return true;"
                
                If DontRenderGroup Then
                    p.HideDelay = 0
                    g(1).Actions.onmouseover.Type = atcNone
                Else
                    g(1).Actions.onmouseover.TargetMenu = 1
                    For i = 1 To UBound(MenuCmds)
                        If MenuCmds(i).parent = idx Then
                            ReDim Preserve c(UBound(c) + 1)
                            c(UBound(c)) = MenuCmds(i)
                            c(UBound(c)).parent = 1
                        End If
                    Next i
                End If
            End If
        Else
            ReDim p.Toolbars(0)
            ReDim c(1)
            c(1) = MenuCmds(idx)
            g(1) = MenuGrps(c(1).parent)
            g(1).Name = "grpp1"
            'g(1).fWidth = GetDivWidth(c(1).Parent)
            c(1).parent = 1
        End If
    End If
    
    For i = 1 To UBound(c)
        With c(i).Actions
            .onclick.Type = atcNone
            .onmouseover.Type = atcNone
        End With
    Next i
    For i = 1 To UBound(g)
        g(i).Actions.onclick.Type = atcNone
        If g(i).Actions.onmouseover.Type = atcNewWindow Or g(i).Actions.onmouseover.Type = atcURL Then g(i).Actions.onmouseover.Type = atcNone
    Next i
    
    If UBound(p.Toolbars) > 0 Then
        g(1).Actions.onmouseover.Type = atcNone
        g(1).Actions.onclick.Type = atcNone
        
        With p.Toolbars(1)
            .Alignment = tbacTopLeft
            .OffsetH = 5
            .OffsetV = 5
            .Condition = "return true;"
        End With
    End If
    
    CompileProject g, c, p, Preferences, params, True, , StatesPath, , , , , True
    If IsTBSel Then
        CopyProjectImages StatesPath, , -2
    Else
        If IsG Then
            CopyProjectImages StatesPath, idx, , -2
        Else
            CopyProjectImages StatesPath, MenuCmds(idx).parent, idx, -2
        End If
    End If
    
    If IsG Then
        If tbi = 0 Then
            sStr = "ShowMenu2('grpp1',5,5,false);function HideMenus(e) {SetPointerPos(e);if(!IsOverMenus()) {ClearTimer(mhdHnd);if(om[1].sc) HoverSel(1);}}"
        End If
    Else
        If Not IsTB Then
            sStr = "ShowMenu2('grpp1',5,5,false);function HideMenus(e) {SetPointerPos(e);if(!IsOverMenus()) {ClearTimer(mhdHnd);if(om[1].sc) HoverSel(1);}}"
        End If
    End If
    If Not IsInIDE Then sStr = sStr + "document.oncontextmenu=new Function(""return false;"")"
    
    If PreviewIsOn Then
        bgColor = Val(frmPreview.cmdColor.tag)
    Else
        bgColor = GetSetting(App.EXEName, "PreviewWinPos", "BackColor", vbWhite)
    End If
    
    If LivePreviewCharset = "_autodetect_all" Then
        sStr = "<html>" + _
            "<meta http-equiv=""X-UA-Compatible"" content=""IE=9"">" + _
            "<body bgcolor=""" + GetRGB(bgColor, True) + """><script language=javascript>" + LoadFile(StatesPath + "menu.js") + sStr + "</script></body></html>"
    Else
        sStr = "<html>" + _
            "<meta http-equiv=""X-UA-Compatible"" content=""IE=9"">" + _
            "<meta http-equiv=""Content-Type"" content=""text/html; charset=" + LivePreviewCharset + """>" + _
            "<body bgcolor=""" + GetRGB(bgColor, True) + """><script language=javascript>" + LoadFile(StatesPath + "menu.js") + sStr + "</script></body></html>"
    End If
            
    If lastHTML = sStr And altWB Is Nothing Then
        LivePreviewIsBusy = False
    Else
        tmrResetLivePreviewBusyState.Enabled = False
        
        lastHTML = sStr
        SaveFile StatesPath + "lpv.html", sStr
        
        If altWB Is Nothing Then
            wbMainPreview.Refresh
        Else
            altWB.Refresh
            lastHTML = ""
        End If
        
        tmrResetLivePreviewBusyState.Enabled = True
    End If
    
ExitSub:
    If Not lastWB Is Nothing Then
        If Not tmrDelayedLivePreview.Enabled Then Set lastWB = Nothing
    End If
        
    If IsInIDE Then
        If Err.number <> 0 Then
            MsgBox "Live Preview Error " & Err.number & ": " + Err.Description
        End If
    End If

End Sub

Friend Sub InitLivePreview(Optional altWB As WebBrowser)

    Static IconsOK As Boolean
    
    If Not IconsOK Then
        cs = GetSysCharsets
        IconsOK = True
        'Set sbLivePreviewBackColor.Picture = ilIcons.ListImages("mnuMenuColor|mnuContextColor").Picture
        'Set sbPreview.Picture = ilIcons.ListImages("mnuToolsPreview").Picture
    End If

    SaveFile StatesPath + "lpv.html", "<html><body></body></html>"
    
    If altWB Is Nothing Then
        wbMainPreview.Navigate StatesPath + "lpv.html"
        DoDelayedLivePreview
    Else
        altWB.Navigate StatesPath + "lpv.html"
        Select Case SelectiveLauncher
            Case slcSelCursor, slcSelFont, slcSelSelFX
                DoLivePreview altWB
            Case Else
                DoLivePreview altWB, IsGroup(tvMenus.SelectedItem.key)
        End Select
    End If
    
End Sub

Private Sub tmrDelayedLivePreview_Timer()

    tmrDelayedLivePreview.Enabled = False
    
    If LivePreviewIsBusy Or IsCompiling Then
        DoDelayedLivePreview
    Else
        DoLivePreview
    End If

End Sub

Private Sub wbMainPreview_NavigateComplete2(ByVal pDisp As Object, url As Variant)

    LivePreviewIsBusy = False

End Sub

Private Sub wbMainPreview_NavigateError(ByVal pDisp As Object, url As Variant, Frame As Variant, StatusCode As Variant, Cancel As Boolean)

    LivePreviewIsBusy = False

End Sub
