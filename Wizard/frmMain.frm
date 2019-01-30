VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{0441225F-21E4-4AB9-94C5-F74C72F390FB}#1.1#0"; "ColorPicker.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DMB Wizard"
   ClientHeight    =   11190
   ClientLeft      =   30
   ClientTop       =   2265
   ClientWidth     =   15045
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11190
   ScaleWidth      =   15045
   Begin VB.PictureBox picTBNumItems 
      Height          =   1845
      Left            =   3435
      ScaleHeight     =   1785
      ScaleWidth      =   5130
      TabIndex        =   14
      Top             =   5145
      Visible         =   0   'False
      Width           =   5190
      Begin MSComCtl2.UpDown udNumGroups 
         Height          =   285
         Left            =   496
         TabIndex        =   63
         Top             =   60
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   5
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtGroups"
         BuddyDispid     =   196610
         OrigLeft        =   765
         OrigTop         =   105
         OrigRight       =   1005
         OrigBottom      =   345
         Max             =   20
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtGroups 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   75
         TabIndex        =   15
         Text            =   "5"
         Top             =   60
         Width           =   420
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Select the number of items that you want to include on your toolbar"
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   75
         TabIndex        =   74
         Top             =   375
         Width           =   4875
      End
   End
   Begin VB.PictureBox picObj1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   2070
      ScaleHeight     =   285
      ScaleWidth      =   420
      TabIndex        =   81
      Top             =   8970
      Visible         =   0   'False
      Width           =   420
   End
   Begin VB.PictureBox picCommands 
      Height          =   3735
      Left            =   7200
      ScaleHeight     =   3675
      ScaleWidth      =   7065
      TabIndex        =   41
      Top             =   7080
      Visible         =   0   'False
      Width           =   7125
      Begin VB.CommandButton cmdApplyFromPreset 
         Caption         =   "Apply Style from Preset..."
         Height          =   375
         Left            =   3930
         TabIndex        =   80
         Top             =   2370
         Width           =   2130
      End
      Begin VB.CommandButton cmdDelete2 
         Caption         =   "Delete"
         Height          =   330
         Left            =   1080
         TabIndex        =   79
         Top             =   2805
         Width           =   975
      End
      Begin MSComCtl2.UpDown udNumCommands 
         Height          =   285
         Left            =   4350
         TabIndex        =   64
         Top             =   300
         Width           =   240
         _ExtentX        =   450
         _ExtentY        =   503
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtNumCmds"
         BuddyDispid     =   196619
         OrigLeft        =   4440
         OrigTop         =   450
         OrigRight       =   4680
         OrigBottom      =   645
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdBrowse 
         Height          =   315
         Left            =   6075
         TabIndex        =   53
         Top             =   1005
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmMain.frx":2CFA
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
      Begin VB.TextBox txtURL 
         Height          =   285
         Left            =   3930
         TabIndex        =   52
         Top             =   1020
         Width           =   2070
      End
      Begin VB.CommandButton cmdRename2 
         Caption         =   "Rename"
         Height          =   330
         Left            =   60
         TabIndex        =   46
         Top             =   2805
         Width           =   975
      End
      Begin VB.CommandButton cmdCreateCmds 
         Caption         =   "Create"
         Height          =   285
         Left            =   4665
         TabIndex        =   45
         Top             =   300
         Width           =   840
      End
      Begin VB.TextBox txtNumCmds 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   3930
         TabIndex        =   44
         Text            =   "0"
         Top             =   300
         Width           =   420
      End
      Begin MSComctlLib.TreeView tvGroups 
         Height          =   2670
         Left            =   60
         TabIndex        =   42
         Top             =   75
         Width           =   3810
         _ExtentX        =   6720
         _ExtentY        =   4710
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LineStyle       =   1
         Style           =   7
         Appearance      =   1
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hyperlink"
         Height          =   195
         Left            =   3930
         TabIndex        =   51
         Top             =   810
         Width           =   660
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Commands"
         Height          =   195
         Left            =   3930
         TabIndex        =   43
         Top             =   75
         Width           =   780
      End
   End
   Begin xfxLine3D.ucLine3D uc3DLine2 
      Height          =   30
      Left            =   0
      TabIndex        =   9
      Top             =   945
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   53
   End
   Begin VB.Frame frameControls 
      Height          =   3390
      Left            =   2505
      TabIndex        =   10
      Top             =   2280
      Width           =   6750
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1440
      Left            =   2490
      ScaleHeight     =   1440
      ScaleWidth      =   45
      TabIndex        =   77
      Top             =   945
      Width           =   45
   End
   Begin VB.PictureBox picTBAlignment 
      Height          =   1665
      Left            =   2400
      ScaleHeight     =   1605
      ScaleWidth      =   2730
      TabIndex        =   11
      Top             =   9930
      Visible         =   0   'False
      Width           =   2790
      Begin VB.OptionButton opTBAlignment 
         Caption         =   "Horizontal"
         Height          =   270
         Index           =   1
         Left            =   45
         TabIndex        =   13
         Top             =   870
         Value           =   -1  'True
         Width           =   1260
      End
      Begin VB.OptionButton opTBAlignment 
         Caption         =   "Vertical"
         Height          =   270
         Index           =   0
         Left            =   45
         TabIndex        =   12
         Top             =   45
         Width           =   1260
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Use this option if you want the menu bar to appear on the top of your pages and horizontally centered."
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   90
         TabIndex        =   76
         Top             =   1110
         Width           =   5325
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Use this option if you want the menu bar to appear aligned to the left and vertically centered."
         ForeColor       =   &H00404040&
         Height          =   405
         Left            =   90
         TabIndex        =   75
         Top             =   300
         Width           =   5325
      End
   End
   Begin VB.PictureBox picIntro 
      Height          =   2550
      Left            =   9480
      ScaleHeight     =   2490
      ScaleWidth      =   5325
      TabIndex        =   65
      Top             =   8925
      Visible         =   0   'False
      Width           =   5385
      Begin VB.TextBox txtRootWeb 
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1470
         Width           =   3120
      End
      Begin xfxLine3D.ucLine3D uc3DLine4 
         Height          =   30
         Left            =   30
         TabIndex        =   68
         Top             =   1050
         Width           =   6675
         _ExtentX        =   11774
         _ExtentY        =   53
      End
      Begin VB.TextBox txtProjName 
         Height          =   285
         Left            =   90
         TabIndex        =   67
         Top             =   360
         Width           =   2415
      End
      Begin SmartButtonProject.SmartButton cmdBrowseRoot 
         Height          =   315
         Left            =   3285
         TabIndex        =   71
         Top             =   1455
         Width           =   360
         _ExtentX        =   635
         _ExtentY        =   556
         Picture         =   "frmMain.frx":2E54
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
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMain.frx":2FAE
         ForeColor       =   &H00404040&
         Height          =   960
         Left            =   105
         TabIndex        =   73
         Top             =   1770
         Width           =   5250
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Use this field to give a name for your project."
         ForeColor       =   &H00404040&
         Height          =   195
         Left            =   90
         TabIndex        =   72
         Top             =   675
         Width           =   3270
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Web Site Location"
         Height          =   195
         Left            =   90
         TabIndex        =   69
         Top             =   1260
         Width           =   1290
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Project Name"
         Height          =   195
         Left            =   90
         TabIndex        =   66
         Top             =   135
         Width           =   960
      End
   End
   Begin VB.PictureBox picPreview 
      Height          =   5205
      Left            =   14205
      ScaleHeight     =   5145
      ScaleWidth      =   6885
      TabIndex        =   58
      Top             =   2865
      Visible         =   0   'False
      Width           =   6945
      Begin xfxLine3D.ucLine3D uc3DLine3 
         Height          =   30
         Left            =   30
         TabIndex        =   61
         Top             =   4635
         Width           =   6795
         _ExtentX        =   11986
         _ExtentY        =   53
      End
      Begin VB.CommandButton cmdClosePreview 
         Caption         =   "Close"
         Height          =   360
         Left            =   6015
         TabIndex        =   60
         Top             =   4695
         Width           =   765
      End
      Begin SHDocVwCtl.WebBrowser wbPreview 
         Height          =   4620
         Left            =   30
         TabIndex        =   59
         Top             =   30
         Width           =   6795
         ExtentX         =   11986
         ExtentY         =   8149
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
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
      Begin VB.Shape Shape1 
         BorderColor     =   &H00808080&
         BorderWidth     =   3
         Height          =   5115
         Left            =   15
         Top             =   0
         Width           =   6840
      End
   End
   Begin VB.CommandButton cmdPreviewInt 
      Caption         =   "Preview"
      Height          =   360
      Left            =   4170
      TabIndex        =   62
      Top             =   5775
      Width           =   1140
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   330
      Top             =   8685
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView tvMenus 
      Height          =   780
      Left            =   270
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   9960
      Visible         =   0   'False
      WhatsThisHelpID =   20000
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   1376
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      Style           =   1
      ImageList       =   "ilIcons"
      Appearance      =   1
      OLEDragMode     =   1
      OLEDropMode     =   1
   End
   Begin VB.PictureBox picItemIcon2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1740
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   56
      Top             =   10260
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picItemIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1440
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   55
      Top             =   10260
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   285
      Top             =   9285
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483634
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   47
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3071
            Key             =   "Color..."
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":31CD
            Key             =   "Font..."
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3329
            Key             =   "Cursor..."
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3485
            Key             =   "Image..."
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":35E1
            Key             =   "Special Effects..."
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":373D
            Key             =   "Move &Up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3899
            Key             =   "Move &Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":39F5
            Key             =   "&New"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3F39
            Key             =   "&Open..."
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":447D
            Key             =   "&Save"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":49C1
            Key             =   "&Copy..."
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4F05
            Key             =   "&Paste..."
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5449
            Key             =   "&Undo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":598D
            Key             =   "&Redo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5ED1
            Key             =   "Re&move"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5FE9
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6101
            Key             =   "Add &Group"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":669D
            Key             =   "Add &Command"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6C39
            Key             =   "Add &Separator"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":71D5
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7771
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D0D
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":82A9
            Key             =   "&Preview..."
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8845
            Key             =   "&HotSpots Editor..."
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8DE1
            Key             =   "Project &Properties..."
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8EF9
            Key             =   "&Compile"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9011
            Key             =   "&AddIn Editor..."
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9465
            Key             =   "&Contents..."
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BC19
            Key             =   "&Upgrade..."
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C1C1
            Key             =   "&FAQ"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C515
            Key             =   "&News"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C869
            Key             =   "&Support"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CBBD
            Key             =   "&Home Page"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CF11
            Key             =   "&Public Forum"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D265
            Key             =   ""
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D801
            Key             =   ""
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DD9D
            Key             =   ""
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DDF9
            Key             =   "&Find..."
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DF0B
            Key             =   "NoEvents"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E4A5
            Key             =   "OverCascade"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EA3F
            Key             =   "Click"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EFD9
            Key             =   "ClickCascade"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F573
            Key             =   "DoubleClick"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FB0D
            Key             =   "DoubleClickCascade"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":100A7
            Key             =   "Over"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10641
            Key             =   "Disabled"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10BDB
            Key             =   "Margins..."
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picRsc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1395
      ScaleHeight     =   270
      ScaleWidth      =   360
      TabIndex        =   54
      Top             =   9870
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   90
      ScaleHeight     =   270
      ScaleWidth      =   3945
      TabIndex        =   50
      Top             =   5820
      Visible         =   0   'False
      Width           =   3945
   End
   Begin VB.PictureBox picLastStep 
      Height          =   3225
      Left            =   10005
      ScaleHeight     =   3165
      ScaleWidth      =   3885
      TabIndex        =   47
      Top             =   4920
      Visible         =   0   'False
      Width           =   3945
      Begin VB.CommandButton cmdSaveProject 
         Caption         =   "Save Project"
         Height          =   465
         Left            =   120
         TabIndex        =   78
         Top             =   1012
         Width           =   2475
      End
      Begin VB.CommandButton cmdInstall 
         Caption         =   "Install Menus"
         Height          =   465
         Left            =   120
         TabIndex        =   49
         Top             =   405
         Width           =   2475
      End
      Begin VB.CommandButton cmdOpenDMB 
         Caption         =   "Open in DHTML Menu Builder"
         Height          =   465
         Left            =   120
         TabIndex        =   48
         Top             =   1620
         Width           =   2475
      End
   End
   Begin VB.PictureBox picCommonFont 
      Height          =   2190
      Left            =   9360
      ScaleHeight     =   2130
      ScaleWidth      =   5325
      TabIndex        =   32
      Top             =   90
      Visible         =   0   'False
      Width           =   5385
      Begin VB.ComboBox cmbFontAlignment 
         Height          =   315
         ItemData        =   "frmMain.frx":10F75
         Left            =   60
         List            =   "frmMain.frx":10F82
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   1710
         Width           =   1140
      End
      Begin VB.Frame Frame2 
         Caption         =   "Normal"
         Height          =   1380
         Left            =   60
         TabIndex        =   36
         Top             =   0
         Width           =   2535
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   300
            Index           =   0
            Left            =   675
            TabIndex        =   37
            Top             =   1005
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            Caption         =   "Change"
            Picture         =   "frmMain.frx":10F9B
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
         Begin VB.Label lblFont 
            BackStyle       =   0  'Transparent
            Height          =   435
            Index           =   0
            Left            =   105
            TabIndex        =   38
            Top             =   315
            Width           =   2265
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mouse Over"
         Height          =   1380
         Left            =   2715
         TabIndex        =   33
         Top             =   0
         Width           =   2535
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   300
            Index           =   1
            Left            =   690
            TabIndex        =   34
            Top             =   1005
            Width           =   1125
            _ExtentX        =   1984
            _ExtentY        =   529
            Caption         =   "Change"
            Picture         =   "frmMain.frx":11335
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
         Begin VB.Label lblFont 
            BackStyle       =   0  'Transparent
            Height          =   435
            Index           =   1
            Left            =   120
            TabIndex        =   35
            Top             =   300
            Width           =   2265
         End
      End
      Begin VB.Label lblAlignment 
         AutoSize        =   -1  'True
         Caption         =   "Alignment"
         Height          =   210
         Left            =   60
         TabIndex        =   39
         Top             =   1485
         Width           =   825
      End
   End
   Begin VB.PictureBox picColorPicker 
      BackColor       =   &H00808080&
      Height          =   3855
      Left            =   9420
      ScaleHeight     =   3795
      ScaleWidth      =   4560
      TabIndex        =   30
      Top             =   1290
      Width           =   4620
      Begin ColorPicker.ucColorPicker ucCP 
         Height          =   3750
         Left            =   15
         TabIndex        =   31
         Top             =   15
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   6615
      End
   End
   Begin VB.PictureBox picCommonColors 
      Height          =   1110
      Left            =   465
      ScaleHeight     =   1050
      ScaleWidth      =   5445
      TabIndex        =   19
      Top             =   6900
      Visible         =   0   'False
      Width           =   5505
      Begin VB.Frame frameNormal 
         Caption         =   "Normal"
         Height          =   1020
         Left            =   60
         TabIndex        =   25
         Top             =   0
         Width           =   2535
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   240
            Index           =   0
            Left            =   1800
            TabIndex        =   26
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
            TabIndex        =   27
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
         Begin VB.Label lblTextColorN 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text Color"
            Height          =   195
            Left            =   855
            TabIndex        =   29
            Top             =   270
            Width           =   750
         End
         Begin VB.Label lblBackColorN 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Back Color"
            Height          =   195
            Left            =   855
            TabIndex        =   28
            Top             =   600
            Width           =   750
         End
      End
      Begin VB.Frame frameHover 
         Caption         =   "Mouse Over"
         Height          =   1020
         Left            =   2700
         TabIndex        =   20
         Top             =   0
         Width           =   2535
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   240
            Index           =   2
            Left            =   1800
            TabIndex        =   21
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
            TabIndex        =   22
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
         Begin VB.Label lblTextColorO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Text Color"
            Height          =   195
            Left            =   855
            TabIndex        =   24
            Top             =   270
            Width           =   750
         End
         Begin VB.Label lblBackColorO 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Back Color"
            Height          =   195
            Left            =   855
            TabIndex        =   23
            Top             =   600
            Width           =   750
         End
      End
   End
   Begin VB.PictureBox picTBGroupsNames 
      Height          =   2835
      Left            =   3930
      ScaleHeight     =   2775
      ScaleWidth      =   3015
      TabIndex        =   16
      Top             =   7380
      Visible         =   0   'False
      Width           =   3075
      Begin VB.CommandButton cmdRename 
         Caption         =   "Rename"
         Height          =   330
         Left            =   45
         TabIndex        =   18
         Top             =   2340
         Width           =   975
      End
      Begin MSComctlLib.ListView lvGroups 
         Height          =   2205
         Left            =   45
         TabIndex        =   17
         Top             =   60
         Width           =   2880
         _ExtentX        =   5080
         _ExtentY        =   3889
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "chName"
            Text            =   "Name"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   -15
      TabIndex        =   8
      Top             =   5640
      Width           =   8895
      _ExtentX        =   15690
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "� Back"
      Height          =   360
      Left            =   5460
      TabIndex        =   7
      Top             =   5775
      Width           =   1140
   End
   Begin VB.CommandButton cmdNext 
      Caption         =   "Next �"
      Height          =   360
      Left            =   6660
      TabIndex        =   6
      Top             =   5775
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   360
      Left            =   7950
      TabIndex        =   5
      Top             =   5775
      Width           =   1140
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   2580
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Text            =   "frmMain.frx":116CF
      Top             =   975
      Width           =   6615
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000009&
      BorderStyle     =   0  'None
      Height          =   945
      Left            =   0
      ScaleHeight     =   945
      ScaleWidth      =   9195
      TabIndex        =   1
      Top             =   0
      Width           =   9195
      Begin VB.Image Image1 
         Height          =   915
         Left            =   7650
         Picture         =   "frmMain.frx":116DE
         Top             =   0
         Width           =   1635
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Caption         =   "Basic info..."
         Height          =   540
         Left            =   210
         TabIndex        =   3
         Top             =   315
         Width           =   7395
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dialog Title"
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
         Left            =   210
         TabIndex        =   2
         Top             =   75
         Width           =   945
      End
   End
   Begin VB.PictureBox picGraphic 
      BorderStyle     =   0  'None
      Height          =   4710
      Left            =   0
      Picture         =   "frmMain.frx":1654A
      ScaleHeight     =   4710
      ScaleWidth      =   2535
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type WizardElement
    Title As String
    BasicInfo As String
    Graphic As Long
    Description As String
    Controls As PictureBox
    CanGoBack As Boolean
    CanGoNext As Boolean
    CanCancel As Boolean
    CanPreview As Boolean
End Type

Dim IsUsingPreset As Boolean

Dim wels() As WizardElement
Dim Index As Integer

Dim SelColorBtn As Integer
Dim CancelMoveOn As Boolean

Private Sub cmdApplyFromPreset_Click()

    frmSelPreset.Show vbModal
    IsUsingPreset = True

End Sub

Private Sub cmdBack_Click()

    SetElementValues
    Index = Index - 1
    GetElementValues
    LoadWizardElement

End Sub

Private Sub cmdBrowse_Click()

    Dim i As Integer

    With cDlg
        .DialogTitle = "Select the document that you want to link to"
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames
        .Filter = SupportedHTMLDocs
        .InitDir = Project.UserConfigs(0).RootWeb
        .FileName = ""
        On Error Resume Next
        .ShowOpen
        On Error GoTo 0
        If Err.number > 0 Or .FileName = "" Then
            Exit Sub
        End If
        txtURL.Text = ConvertPath(.FileName)
        SetURL
    End With

End Sub

Private Sub SetURL()

    If tvGroups.SelectedItem.Parent Is Nothing Then
        With MenuGrps(tvGroups.SelectedItem.Tag)
            If txtURL.Text = "" Then
                .Actions.onclick.Type = atcNone
            Else
                .Actions.onclick.Type = atcURL
                .Actions.onclick.TargetFrame = "_self"
            End If
            .Actions.onclick.url = txtURL.Text
        End With
    Else
        With MenuCmds(tvGroups.SelectedItem.Tag)
            If txtURL.Text = "" Then
                .Actions.onclick.Type = atcNone
            Else
                .Actions.onclick.Type = atcURL
                .Actions.onclick.TargetFrame = "_self"
            End If
            .Actions.onclick.url = txtURL.Text
        End With
    End If

End Sub

Private Sub cmdBrowseRoot_Click()

    Dim Path As String
    
    On Error Resume Next
    
    Me.Enabled = False
    
    Path = txtRootWeb.Text
    If Dir(Path) = "" Or Err.number Then
        Path = ""
    Else
        Path = UnqualifyPath(Path)
    End If
    Path = BrowseForFolderByPath(Path, "Select the folder where the local copy of your web site is stored", Me)
    
    If Path <> "" Then txtRootWeb.Text = AddTrailingSlash(Path, "\")
    
    Me.Enabled = True
    Me.SetFocus

End Sub

Private Sub cmdCancel_Click()

    If cmdCancel.Caption <> "Finish" Then
        If MsgBox("Are you sure you want to cancel this Wizard?", vbQuestion + vbYesNo, "Close Wizard?") = vbNo Then Exit Sub
    Else
        SaveMenu
    End If
    Unload Me

End Sub

Private Sub cmdChange_Click(Index As Integer)

    With lblFont(Index)
        SelFont.Name = .FontName
        SelFont.Size = pt2px(.FontSize)
        SelFont.Bold = .FontBold
        SelFont.Italic = .FontItalic
        SelFont.Underline = .FontUnderline
        SelFont.IsSubst = False
    End With

    With frmFontDialog
        .Show vbModal
        .Move (Screen.Width - .Width) \ 2, (Screen.Height - .Height) \ 2
        Do
            DoEvents
        Loop While .Visible
    End With
    
    If SelFont.IsValid Then
        With lblFont(Index)
            .FontName = SelFont.Name
            .FontSize = px2pt(SelFont.Size)
            .FontBold = SelFont.Bold
            .FontItalic = SelFont.Italic
            .FontUnderline = SelFont.Underline
        End With
        UpdateFontLabel Index
    End If

End Sub

Private Sub cmdClosePreview_Click()

    picPreview.Visible = False
    RestoreButtonsState

End Sub

Private Sub cmdColor_Click(Index As Integer)

    SaveButtonsState
    
    SelColorBtn = Index

    With picColorPicker
        .Move (Width - .Width) \ 2, (Height - .Height) \ 2
        .ZOrder 0
        .Visible = True
        SetColor ucCP.Color, cmdColor(Index)
    End With
    
    cmdColor(0).Enabled = False
    cmdColor(1).Enabled = False
    cmdColor(2).Enabled = False
    cmdColor(3).Enabled = False

End Sub

Private Sub cmdCreateCmds_Click()

    Dim i As Integer
    Dim k As Integer
    Dim ParentGroup As Integer
    Dim SelGrp As Node
    Dim tNode As Node
    
    If Not IsGroup(tvGroups.SelectedItem) Then CreateSubMenu
    
    Set SelGrp = tvGroups.SelectedItem
    If SelGrp Is Nothing Then Set SelGrp = tvGroups.Nodes(1)
    ParentGroup = SelGrp.Tag
    
    DeleteCommands ParentGroup
    
    ReDim Preserve MenuCmds(UBound(MenuCmds) + Val(txtNumCmds.Text))
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            If .Parent = 0 Then
                If k = 0 Then k = i
                .Parent = ParentGroup
                .Caption = "Command " & (i - k) + 1
                
                If .NormalFont.FontName = "" Then
                    .CmdsFXhColor = MenuGrps(.Parent).CmdsFXhColor
                    .CmdsFXnColor = MenuGrps(.Parent).CmdsFXnColor
                    .CmdsFXNormal = MenuGrps(.Parent).CmdsFXNormal
                    .CmdsFXOver = MenuGrps(.Parent).CmdsFXOver
                    .CmdsFXSize = MenuGrps(.Parent).CmdsFXSize
                    .CmdsMarginX = MenuGrps(.Parent).CmdsMarginX
                    .CmdsMarginY = MenuGrps(.Parent).CmdsMarginY
                    .hBackColor = MenuGrps(.Parent).hBackColor
                    .HoverFont = MenuGrps(.Parent).DefHoverFont
                    .hTextColor = MenuGrps(.Parent).hTextColor
                    .iCursor = MenuGrps(.Parent).iCursor
                    .nBackColor = MenuGrps(.Parent).nBackColor
                    .NormalFont = MenuGrps(.Parent).DefNormalFont
                    .nTextColor = MenuGrps(.Parent).nTextColor
                    
                    .LeftImage = MenuGrps(.Parent).tbiLeftImage
                    .BackImage = MenuGrps(.Parent).tbiBackImage
                    .RightImage = MenuGrps(.Parent).tbiRightImage
                End If
            End If
        End With
    Next i
    
    With MenuGrps(ParentGroup).Actions.onmouseover
        .Type = atcCascade
        .TargetMenu = ParentGroup
    End With
    
    GetElementValues
    
    For Each tNode In tvGroups.Nodes
        If tNode.Text = SelGrp.Text Then
            tNode.EnsureVisible
            tNode.Selected = True
            Exit For
        End If
    Next tNode
    
End Sub

Private Sub CreateSubMenu()

    Dim ni As Integer
    Dim nItem As Node
    Dim scID As Integer
    
    ni = UBound(MenuGrps) + 1

    ReDim Preserve MenuGrps(ni)
    SetGroupDefaults ni
    MenuGrps(ni).IncludeInToolbar = False
    
    With MenuCmds(tvGroups.SelectedItem.Tag)
        If Right(.Caption, 1) <> "�" Then .Caption = .Caption + " �"
        With .Actions.onmouseover
            .Type = atcCascade
            .TargetMenu = ni
            .TargetMenuAlignment = gacRightTop
        End With
        MenuGrps(ni).CmdsFXhColor = MenuGrps(.Parent).CmdsFXhColor
        MenuGrps(ni).CmdsFXnColor = MenuGrps(.Parent).CmdsFXnColor
        MenuGrps(ni).CmdsFXNormal = MenuGrps(.Parent).CmdsFXNormal
        MenuGrps(ni).CmdsFXOver = MenuGrps(.Parent).CmdsFXOver
        MenuGrps(ni).CmdsFXSize = MenuGrps(.Parent).CmdsFXSize
        MenuGrps(ni).CmdsMarginX = MenuGrps(.Parent).CmdsMarginX
        MenuGrps(ni).CmdsMarginY = MenuGrps(.Parent).CmdsMarginY
        MenuGrps(ni).hBackColor = MenuGrps(.Parent).hBackColor
        MenuGrps(ni).DefHoverFont = MenuGrps(.Parent).DefHoverFont
        MenuGrps(ni).hTextColor = MenuGrps(.Parent).hTextColor
        MenuGrps(ni).iCursor = MenuGrps(.Parent).iCursor
        MenuGrps(ni).nBackColor = MenuGrps(.Parent).nBackColor
        MenuGrps(ni).DefNormalFont = MenuGrps(.Parent).DefNormalFont
        MenuGrps(ni).nTextColor = MenuGrps(.Parent).nTextColor
        
        MenuGrps(ni).CornersImages = MenuGrps(.Parent).CornersImages
        MenuGrps(ni).ContentsMarginH = MenuGrps(.Parent).ContentsMarginH
        MenuGrps(ni).ContentsMarginV = MenuGrps(.Parent).ContentsMarginV
        
        MenuGrps(ni).tbiLeftImage = MenuGrps(.Parent).tbiLeftImage
        MenuGrps(ni).tbiBackImage = MenuGrps(.Parent).tbiBackImage
        MenuGrps(ni).tbiRightImage = MenuGrps(.Parent).tbiRightImage
        
        MenuGrps(ni).Compile = True
        .Compile = True
    End With
    
    Set nItem = tvGroups.Nodes.Add(tvGroups.SelectedItem, tvwChild, , MenuGrps(ni).Caption)
    nItem.Tag = ni
    nItem.Bold = True
    
    nItem.Parent.Expanded = True
    nItem.Selected = True

End Sub

Private Sub cmdDelete2_Click()

    If Not tvGroups.SelectedItem Is Nothing Then DeleteItem

End Sub

Private Sub cmdInstall_Click()

    Dim AddInP() As AddInParameter

    GlobalizeSettings
    
    On Error Resume Next
    
    With Project.UserConfigs(0)
        MkDir .CompiledPath
        MkDir .ImagesPath
    End With
    
    CompileProject MenuGrps, MenuCmds, Project, Preferences, AddInP
    With Project
        .DOMCode = ""
        .DOMFramesCode = ""
        .NSCode = ""
        .NSFramesCode = ""
    End With
    
    frmInstallMenus.Show vbModal

End Sub

Private Sub cmdNext_Click()

    SetElementValues
    If Not CancelMoveOn Then
        Index = Index + 1
        LoadWizardElement
        GetElementValues
    End If
    CancelMoveOn = False

End Sub

Private Sub cmdOpenDMB_Click()

    GlobalizeSettings
    SaveMenu
    If FileExists(Project.FileName) Then
        RunShellExecute "open", GetFileName(Project.FileName), "", GetFilePath(Project.FileName), 1
    End If
    
    cmdCancel_Click
    
End Sub

Private Sub SaveMenu()
    
    Dim ff As Integer
    
    If Not FileExists(Project.FileName) Then
        On Error Resume Next
        With cDlg
            .DialogTitle = "Save Project"
            .FileName = Project.FileName
            .Filter = "DHTML Menu Builder Projects|*.dmb"
            .Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
            .ShowSave
            If Err.number > 0 Or .FileName = "" Then
                Exit Sub
            End If
            Project.FileName = .FileName
            Project.Name = GetFileName(.FileName)
            Project.Name = Left(Project.Name, Len(Project.Name) - 4)
            Project.HasChanged = True
        End With
        On Error GoTo 0
    End If
    
    FloodPanel.Caption = "Saving"
    
    SaveProject False
    SaveImages ""
    
'    ff = FreeFile
'    Open Project.FileName For Append As ff
'        Print #ff, "[RSC]"
'    Close #ff
    
    FloodPanel.Value = 0
    
End Sub

Private Sub GenPreview()

    Dim AddInP() As AddInParameter
    Dim ff As Integer
    
    GlobalizeSettings
    
    On Error Resume Next
        MkDir AppPath + "Preview\"
    On Error GoTo 0

    CopyProjectImages AppPath + "Preview\"
    CompileProject MenuGrps, MenuCmds, Project, Preferences, AddInP, True, False, AppPath + "Preview\"
    
    ff = FreeFile
    Open AppPath + "Preview\index.html" For Output As #ff
        Print #ff, "<head><title>Wizard Preview</title></head>"
        Print #ff, "<body>"
        Print #ff, "<script language=JavaScript src=menu.js></script>"
        Print #ff, "</body>"
    Close #ff

End Sub

Private Sub cmdPreviewInt_Click()

    SaveButtonsState
    
    SetElementValues
    GenPreview
    
    With picPreview
        .Move (Width - .Width) \ 2, (Height - .Height) \ 2
        .ZOrder 0
        .Visible = True
        wbPreview.Navigate AppPath + "Preview\index.html"
    End With

End Sub

Private Sub cmdRename_Click()

    If Not lvGroups.SelectedItem Is Nothing Then
        lvGroups.SetFocus
        lvGroups.StartLabelEdit
    End If

End Sub

Private Sub cmdRename2_Click()

    If Not tvGroups.SelectedItem Is Nothing Then
        tvGroups.SetFocus
        tvGroups.StartLabelEdit
    End If

End Sub

Private Sub cmdSaveProject_Click()

    SaveMenu

End Sub

Private Sub Form_Load()

    AppPath = GetSetting("DMB", "RegInfo", "InstallPath", AddTrailingSlash(App.Path, "\"))
    TempPath = GetTEMPPath
    StatesPath = TempPath + "States\Default\"
    MkDir2 StatesPath
    cSep = Chr(255) + Chr(255)
    
    Set FloodPanel.PictureControl = picFlood

    Dim UIObjects(1 To 3) As Object
    Set UIObjects(1) = frmMain
    Set UIObjects(2) = FloodPanel
    Set UIObjects(3) = Me
    SetUI UIObjects
    
    Dim VarObjects(1 To 7) As Variant
    VarObjects(1) = AppPath
    VarObjects(2) = ""
    VarObjects(3) = GenLicense
    VarObjects(4) = TempPath
    VarObjects(5) = cSep
    VarObjects(6) = ""
    VarObjects(7) = StatesPath
    SetVars VarObjects

    With Screen
        Width = 9000 + (15 * 15) + GetClientLeft(hWnd)
        Height = 6255 + GetClientTop(hWnd)
        Move (.Width - Width) \ 2, (.Height - Height) \ 2, Width, Height
    End With
    
    Caption = "DHTML Menu Builder Wizard " & App.Major & "." & App.Minor
    
    ReDim MenuGrps(0)
    ReDim MenuCmds(0)
    
    With Preferences
        .SepHeight = GetSetting("DMB", "Preferences", "SepHeight", 11)
        .ImgSpace = Val(GetSetting("DMB", "Preferences", "ImgSpace", 7))
        .language = "eng"
    End With
    
    DoUNICODE = (GetSetting("DMB", "Preferences", "DoUNICODE", 1) = 1)
    
    With Project
        .Name = "Wizard"
        .AbsPath = ""
        .FileName = ""
        .AddIn.Name = ""
        .CodeOptimization = cocAggressive
        .RemoveImageAutoPosCode = True
        .FX = 0
        .HasChanged = False
        
        ReDim .UserConfigs(0)
        With .UserConfigs(0)
            .Name = "Default"
            .Description = "Default configuration"
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
        With .Toolbars(1)
            .Name = "Default"
            .Alignment = tbacTopCenter
            .BackColor = &H0
            .border = 2
            
            .CustX = 0
            .CustY = 0
            
            .FollowHScroll = False
            .FollowVScroll = False
            Erase .Groups
            .OffsetH = 0
            .OffsetV = 0
            .Image = ""
            .JustifyHotSpots = False
            .OffsetH = 0
            .OffsetV = 0
            .Spanning = tscAuto
            .Style = tscHorizonal
            .Separation = 1
            .Compile = True
        End With
        
        .JSFileName = "menu"
        
        .GenDynAPI = False
        .CompileIECode = True
        .CompileNSCode = True
        .CompilehRefFile = True
        
        .MenusOffset.RootMenusX = 0
        .MenusOffset.RootMenusY = 0
        .MenusOffset.SubMenusX = 0
        .MenusOffset.SubMenusY = 0
        
        .UnfoldingSound.onmouseover = ""
        
        .FontSubstitutions = ""
        
        .DoFormsTweak = True
        .DWSupport = False
        .NS4ClipBug = False
        .OPHelperFunctions = False
        .ImageReadySupport = False
        
        .HideDelay = 200
        .SubMenusDelay = 150
        .RootMenusDelay = 15
        .AnimSpeed = 35
        
        .DXFilter = ""
        
        .ExportHTMLParams = GenExpHTMLPref("", Project.Name, Project.FileName)
    End With
    
    InitWizardElements
    StartWizard
    
    If Not FileExists(AppPath + "dmb.exe") Then
        MsgBox "DHTML Menu Builder must be installed in order to run this Wizard", vbCritical + vbOKOnly, "DHTML Menu Builder not detected"
        End
    End If

End Sub

Private Sub InitWizardElements()

    Dim PreviewInfo As String

    ReDim wels(0 To 7)
    
    PreviewInfo = vbCrLf + vbCrLf + "Click the Preview button at any time to see a real sample of how your menus look like."
    
    With wels(0)
        .Title = "Welcome to the DHTML Menu Builder Wizard"
        .BasicInfo = "Creating menus has never been easier!"
        .Description = "This wizard will guide you through all the steps to create your first menu and implement it on your web site within minutes!" + vbCrLf + vbCrLf + "Fill in the two parameters below and when done click on the Next button to start the wizard."
        Set .Controls = picIntro
        .CanCancel = True
        .CanGoBack = False
        .CanGoNext = True
        .CanPreview = False
        .Graphic = 0
    End With
    
    With wels(1)
        .Title = "Set the Toolbar Style"
        .BasicInfo = "Select how you want to display your toolbar"
        .Description = "Toolbars in DHTML Menu Builder can be aligned either vertically or horizontally." + vbCrLf + vbCrLf + "Select how you would want your toolbar to appear on your pages and click Next to continue."
        Set .Controls = picTBAlignment
        .CanCancel = True
        .CanGoBack = True
        .CanGoNext = True
        .CanPreview = False
        .Graphic = 0
    End With
    
    With wels(2)
        .Title = "Toolbar Items"
        .BasicInfo = "How many items do you want to display on your toolbar"
        .Description = "Type the number of items you want to display on your toolbar. These will be the headings or main categories of your menus." + vbCrLf + vbCrLf + "Click Next when you're done."
        Set .Controls = picTBNumItems
        .CanCancel = True
        .CanGoBack = True
        .CanGoNext = True
        .CanPreview = False
        .Graphic = 0
    End With
    
    With wels(3)
        .Title = "Define the Items' Caption"
        .BasicInfo = "Select the caption for each one of your toolbar items"
        .Description = "Click on each and every item to rename it and set the caption to the text that you want each item to display on the toolbar." + PreviewInfo
        Set .Controls = picTBGroupsNames
        .CanCancel = True
        .CanGoBack = True
        .CanGoNext = True
        .CanPreview = True
        .Graphic = 0
    End With
    
    With wels(4)
        .Title = "Define the Items' Colors"
        .BasicInfo = "Select the colors for your toolbar items"
        .Description = "The normal colors define how your toolbar items will look like when the mouse is not over them. While the over colors control how the toolbar items will look like when the mouse moves over them." + PreviewInfo
        Set .Controls = picCommonColors
        .CanCancel = True
        .CanGoBack = True
        .CanGoNext = True
        .CanPreview = True
        .Graphic = 0
    End With
    
    With wels(5)
        .Title = "Define Items' Font"
        .BasicInfo = "Select the font style for your toolbar items"
        .Description = "The normal font style defines how your toolbar items will look like when the mouse is not over them. While the over font style controls how the toolbar items will look like when the mouse moves over them." + PreviewInfo
        Set .Controls = picCommonFont
        .CanCancel = True
        .CanGoBack = True
        .CanGoNext = True
        .CanPreview = True
        .Graphic = 0
    End With
    
    With wels(6)
        .Title = "Create the Menus"
        .BasicInfo = "Create the menus for each one of the toolbar items"
        .Description = "Select each group and define how many items you want to display when selected." + vbCrLf + "Then set the caption for each one of the items." + PreviewInfo
        Set .Controls = picCommands
        .CanCancel = True
        .CanGoBack = True
        .CanGoNext = True
        .CanPreview = True
        .Graphic = 0
    End With
    
    With wels(7)
        .Title = "Your are done!"
        .BasicInfo = "Your new menus are ready"
        .Description = "The wizard has enough information to generate the menus so you can install them on your web site." + vbCrLf + vbCrLf + "Click the 'Install Menus' button to install the menus on your web site at '%%ROOTWEB%%'." + vbCrLf + vbCrLf + "Or click on the 'Open in DHTML Menu Builder' button to open your project in DHTML Menu Builder so you can take advantage of the dozens of styling options offered by the full application."
        Set .Controls = picLastStep
        .CanCancel = True
        .CanGoBack = True
        .CanGoNext = False
        .CanPreview = True
        .Graphic = 0
    End With
    
End Sub

Private Sub StartWizard()

    Index = 0
    LoadWizardElement

End Sub

Private Sub LoadWizardElement()

    Static LastControlsObj As PictureBox

    With wels(Index)
        lblTitle.Caption = .Title
        lblInfo.Caption = .BasicInfo
        txtDesc.Text = Replace(.Description, "%%ROOTWEB%%", Project.UserConfigs(0).RootWeb)
        cmdNext.Enabled = .CanGoNext And Index < UBound(wels)
        cmdBack.Enabled = .CanGoBack And Index > 0
        cmdCancel.Enabled = .CanCancel
        cmdPreviewInt.Enabled = .CanPreview
        cmdCancel.Caption = IIf(Index = UBound(wels), "Finish", "Cancel")
        If Not LastControlsObj Is Nothing Then
            LastControlsObj.Visible = False
        End If
        If Not .Controls Is Nothing Then
            With .Controls
                .BorderStyle = 0
                .Move frameControls.Left + 30, frameControls.Top + 120, frameControls.Width - 60, frameControls.Height - 240
                .Visible = True
                .ZOrder 0
            End With
            Set LastControlsObj = .Controls
        End If
        If .Graphic <> 0 Then
            picGraphic.Picture = LoadResPicture(.Graphic, vbResBitmap)
        End If
    End With

End Sub

Private Sub DeleteGroup(Group As Integer)

    Dim g As Integer
    Dim c As Integer

    DeleteCommands Group

    For g = Group To UBound(MenuGrps) - 1
        MenuGrps(g) = MenuGrps(g + 1)
        With MenuGrps(g).Actions.onmouseover
            If .Type = atcCascade Then .TargetMenu = g
        End With
        
        For c = 1 To UBound(MenuCmds)
            With MenuCmds(c)
                If .Parent = g + 1 Then .Parent = g
            End With
        Next c
    Next g
    ReDim Preserve MenuGrps(UBound(MenuGrps) - 1)
    
    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c)
            If .Actions.onmouseover.Type = atcCascade And .Actions.onmouseover.TargetMenu = Group Then
                .Actions.onmouseover.Type = atcNone
                If Right(.Caption, 2) = " �" Then .Caption = Left(.Caption, Len(.Caption) - 2)
            End If
        End With
    Next c

End Sub

Friend Sub DeleteItem()

    Dim sItem As String
    Dim n As Node
    
    If tvGroups.SelectedItem.Parent Is Nothing Then
        If UBound(MenuGrps) = 1 Then
            MsgBox "The project must contain at least one toolbar item", vbInformation + vbOKOnly, "Error Deleting Item"
            Exit Sub
        End If
    End If
    
    If tvGroups.SelectedItem.Next Is Nothing Then
        If tvGroups.SelectedItem.Previous Is Nothing Then
            ' Do nothing
        Else
            sItem = tvGroups.SelectedItem.Previous.Text
        End If
    Else
        sItem = tvGroups.SelectedItem.Next.Text
    End If

    If IsGroup(tvGroups.SelectedItem) Then
        DeleteGroup Val(tvGroups.SelectedItem.Tag)
    Else
        DeleteCommand Val(tvGroups.SelectedItem.Tag)
    End If
    
    GetElementValues
    
    If sItem <> "" Then
        For Each n In tvGroups.Nodes
            If n.Text = sItem Then
                n.EnsureVisible
                n.Selected = True
                Exit For
            End If
        Next n
    End If
    
    tvGroups.SetFocus

End Sub

Private Sub DeleteCommand(c As Integer)

    For c = c To UBound(MenuCmds) - 1
        MenuCmds(c) = MenuCmds(c + 1)
    Next c
    ReDim Preserve MenuCmds(UBound(MenuCmds) - 1)

End Sub

Private Sub DeleteCommands(ParentGroup As Integer)

    Dim c As Integer
    Dim cc As Integer
    Dim tm As Integer

ReStart:
    For c = 1 To UBound(MenuCmds)
        If MenuCmds(c).Parent = ParentGroup Then
            If MenuCmds(c).Actions.onmouseover.Type = atcCascade Then
                tm = MenuCmds(c).Actions.onmouseover.TargetMenu
                DeleteGroup tm
            End If
            For cc = c To UBound(MenuCmds) - 1
                MenuCmds(cc) = MenuCmds(cc + 1)
            Next cc
            ReDim Preserve MenuCmds(UBound(MenuCmds) - 1)
            GoTo ReStart
        End If
    Next c

End Sub

Private Sub SetElementValues()

          Dim i As Integer
          Dim nItem As ListItem
          Dim tNode As Node
          
10        On Error GoTo ReportError

20        Project.HasChanged = True
30        Select Case Index
              Case 0
40                Project.Name = txtProjName.Text
50                If txtRootWeb.Text = "" Then
60                    CancelMoveOn = True
70                    MsgBox "The 'Web Site Location' parameter cannot be empty", vbInformation + vbOKOnly, "Missing or invalid parameter"
80                Else
90                    Project.UserConfigs(0).RootWeb = txtRootWeb.Text
100                   Project.UserConfigs(0).CompiledPath = txtRootWeb.Text + "menus\"
110                   Project.UserConfigs(0).ImagesPath = txtRootWeb.Text + "menus\images\"
120               End If
130           Case 1
140               Project.Toolbars(1).Style = IIf(opTBAlignment(0).Value, tscVertical, tscHorizonal)
150           Case 2
160               If Val(txtGroups.Text) <= 0 Then
170                   CancelMoveOn = True
180                   MsgBox "The number of groups to create must be greater than 0", vbInformation + vbOKOnly, "Missing or invalid parameter"
190               Else
200                   ReDim Preserve MenuGrps(Val(txtGroups.Text))
210                   For i = 1 To UBound(MenuGrps)
220                       DeleteCommands i
230                       SetGroupDefaults i
240                   Next i
250               End If
260           Case 3
270               For Each nItem In lvGroups.ListItems
280                   MenuGrps(Val(nItem.Tag)).Caption = nItem.Text
290               Next nItem
300           Case 4
310               For i = 1 To UBound(MenuGrps)
320                   MenuGrps(i).nTextColor = cmdColor(0).BackColor
330                   MenuGrps(i).nBackColor = cmdColor(1).BackColor
340                   MenuGrps(i).hTextColor = cmdColor(2).BackColor
350                   MenuGrps(i).hBackColor = cmdColor(3).BackColor
360               Next i
370           Case 5
380               IsUsingPreset = False
390               For i = 1 To UBound(MenuGrps)
400                   With MenuGrps(i).DefNormalFont
410                       .FontName = lblFont(0).FontName
420                       .FontSize = pt2px(lblFont(0).FontSize)
430                       .FontBold = lblFont(0).FontBold
440                       .FontItalic = lblFont(0).FontItalic
450                       .FontUnderline = lblFont(0).FontUnderline
460                   End With
470                   With MenuGrps(i).DefHoverFont
480                       .FontName = lblFont(1).FontName
490                       .FontSize = pt2px(lblFont(1).FontSize)
500                       .FontBold = lblFont(1).FontBold
510                       .FontItalic = lblFont(1).FontItalic
520                       .FontUnderline = lblFont(1).FontUnderline
530                   End With
540                   MenuGrps(i).CaptionAlignment = cmbFontAlignment.ListIndex
550               Next i
560           Case 6
570               For Each tNode In tvGroups.Nodes
580                   If IsGroup(tNode) Then
590                       MenuGrps(tNode.Tag).Caption = tNode.Text
600                   Else
610                       MenuCmds(tNode.Tag).Caption = tNode.Text
620                   End If
630               Next tNode
640       End Select

650       Exit Sub

ReportError:

660       MsgBox "SetElementValues has generated an error " & Err.number & ": " & vbCrLf & _
            Err.Description & " at line " & Erl, vbCritical + vbOKOnly, "Unexpected Error"

End Sub

Private Sub SetGroupDefaults(g As Integer)

    If g > UBound(MenuGrps) Then Exit Sub

    With MenuGrps(g)
        If .Caption = "" Then
            .Caption = "Item " & g
            .nTextColor = vbBlue
            .nBackColor = &HC0C0C0
            .hTextColor = vbWhite
            .hBackColor = vbBlue
            With .DefNormalFont
                .FontName = "Tahoma"
                .FontBold = False
                .FontItalic = False
                .FontSize = 11
                .FontUnderline = False
            End With
            With .DefHoverFont
                .FontName = "Tahoma"
                .FontSize = 11
                .FontBold = False
                .FontItalic = False
                .FontUnderline = False
            End With
            .CaptionAlignment = tacCenter
            .IncludeInToolbar = True
            .Compile = True
        End If
    End With

End Sub

Public Sub GlobalizeSettings()

    Dim c As Integer
    Dim g As Integer
    Dim HasCommands As Boolean
    
    For g = 1 To UBound(MenuGrps)
        With MenuGrps(g)
        .Compile = True
            With .Actions
                .OnDoubleClick.Type = atcNone
                HasCommands = False
                For c = 1 To UBound(MenuCmds)
                    If MenuCmds(c).Parent = g Then
                        HasCommands = True
                        Exit For
                    End If
                Next c
                .onmouseover.Type = IIf(HasCommands, atcCascade, atcNone)
                .onmouseover.TargetMenu = g
            End With
            .Alignment = IIf(Project.Toolbars(1).Style = tscHorizonal, gacBottomLeft, gacRightTop)
            .AlignmentStyle = ascVertical
            .Name = "grp" & Format(g, "00")
            .WinStatus = .Caption
            
            If Not IsUsingPreset Then
                .CaptionAlignment = MenuGrps(1).CaptionAlignment
                .CmdsFXhColor = -2
                .CmdsFXnColor = -2
                .CmdsFXNormal = cfxcNone
                .CmdsFXOver = cfxcNone
                .CmdsFXSize = 1
                .CmdsMarginX = 4
                .CmdsMarginY = 2
                .ContentsMarginH = 2
                .ContentsMarginV = 2
                .Corners.bottomCorner = MenuGrps(1).hBackColor
                .Corners.leftCorner = MenuGrps(1).hBackColor
                .Corners.rightCorner = MenuGrps(1).hBackColor
                .Corners.topCorner = MenuGrps(1).hBackColor
                .DefHoverFont = MenuGrps(1).DefHoverFont
                .DefNormalFont = MenuGrps(1).DefNormalFont
                .nBackColor = MenuGrps(1).nBackColor
                .nTextColor = MenuGrps(1).nTextColor
                .hBackColor = MenuGrps(1).hBackColor
                .hTextColor = MenuGrps(1).hTextColor
                .disabled = False
                .DropShadowColor = &H999999
                .DropShadowSize = 0
                .fHeight = 0
                .frameBorder = 1
                .fWidth = 0
                .iCursor.cType = iccHand
                .iCursor.cFile = ""
                .IsContext = False
                .IsTemplate = False
                .Leading = 1
                .Transparency = 0
                .bColor = .nBackColor
            End If
        End With
        For c = 1 To UBound(MenuCmds)
            With MenuCmds(c)
                .Compile = True
                If .Parent = g Then
                    With .Actions
                        .onclick.Type = atcURL
                        .OnDoubleClick.Type = atcNone
                    End With
                    .Name = "cmd" & MenuGrps(.Parent).Name & Format(c, "00")
                    .WinStatus = .Caption
                    
                    If Not IsUsingPreset Then
                        .Alignment = tacLeft
                        .disabled = False
                        .nTextColor = MenuGrps(.Parent).nTextColor
                        .hTextColor = MenuGrps(.Parent).hTextColor
                        .nBackColor = MenuGrps(.Parent).nBackColor
                        .hBackColor = MenuGrps(.Parent).hBackColor
                        .NormalFont = MenuGrps(.Parent).DefNormalFont
                        .HoverFont = MenuGrps(.Parent).DefHoverFont
                        .iCursor.cType = MenuGrps(.Parent).iCursor.cType
                        .iCursor.cFile = ""
                        .CmdsFXhColor = -2
                        .CmdsFXnColor = -2
                        .CmdsFXNormal = cfxcNone
                        .CmdsFXOver = cfxcNone
                        .CmdsFXSize = 1
                        .CmdsMarginX = 4
                        .CmdsMarginY = 2
                    End If
                End If
            End With
        Next c
    Next g
    
    With Project.Toolbars(1)
        If Not IsUsingPreset Then
            .BackColor = MenuGrps(1).hBackColor
            .border = 1
            .BorderColor = MenuGrps(1).hBackColor
            .ContentsMarginH = 0
            .ContentsMarginV = 0
            .JustifyHotSpots = True
            .Separation = 1
            .Spanning = tscAuto
            .Alignment = IIf(.Style = tscHorizonal, tbacTopCenter, tbacCenterLeft)
            .Compile = True
        End If
        
        ReDim .Groups(0)
        For g = 1 To UBound(MenuGrps)
            If MenuGrps(g).IncludeInToolbar Then
                ReDim Preserve .Groups(UBound(.Groups) + 1)
                .Groups(UBound(.Groups)) = MenuGrps(g).Name
            End If
        Next g
    End With
    
    For c = 1 To UBound(MenuCmds)
        With MenuCmds(c).Actions.onmouseover
            If .Type = atcCascade Then
                MenuGrps(.TargetMenu).IncludeInToolbar = False
            End If
        End With
    Next c

End Sub

Private Sub GetElementValues()

          Dim i As Integer
          Dim j As Integer
          Dim nItem As ListItem
          Dim tItem As Node
          Dim tcItem As Node
          
10        On Error GoTo ReportError

20        Select Case Index
              Case 0
30                txtProjName.Text = Project.Name
40            Case 1
50                opTBAlignment(0).Value = (Project.Toolbars(1).Style = tscVertical)
60                opTBAlignment(1).Value = (Project.Toolbars(1).Style = tscHorizonal)
70            Case 2
80                txtGroups.Text = IIf(UBound(MenuGrps) = 0, 5, UBound(MenuGrps))
90            Case 3
100               With lvGroups.ListItems
110                   .Clear
120                   For i = 1 To UBound(MenuGrps)
130                       Set nItem = .Add(, , MenuGrps(i).Caption)
140                       nItem.Tag = i
150                   Next i
160               End With
170               lvGroups.ColumnHeaders(1).Width = lvGroups.Width - 120
180           Case 4
190               SetColor MenuGrps(1).nTextColor, cmdColor(0)
200               SetColor MenuGrps(1).nBackColor, cmdColor(1)
210               SetColor MenuGrps(1).hTextColor, cmdColor(2)
220               SetColor MenuGrps(1).hBackColor, cmdColor(3)
230           Case 5
240               With lblFont(0)
250                   .FontName = MenuGrps(1).DefNormalFont.FontName
260                   .FontSize = px2pt(MenuGrps(1).DefNormalFont.FontSize)
270                   .FontBold = MenuGrps(1).DefNormalFont.FontBold
280                   .FontItalic = MenuGrps(1).DefNormalFont.FontItalic
290                   .FontUnderline = MenuGrps(1).DefNormalFont.FontUnderline
300               End With
310               With lblFont(1)
320                   .FontName = MenuGrps(1).DefHoverFont.FontName
330                   .FontSize = px2pt(MenuGrps(1).DefHoverFont.FontSize)
340                   .FontBold = MenuGrps(1).DefHoverFont.FontBold
350                   .FontItalic = MenuGrps(1).DefHoverFont.FontItalic
360                   .FontUnderline = MenuGrps(1).DefHoverFont.FontUnderline
370               End With
380               cmbFontAlignment.ListIndex = MenuGrps(1).CaptionAlignment
390               UpdateFontLabel 0
400               UpdateFontLabel 1
410           Case 6
420               With tvGroups.Nodes
430                   .Clear
440                   For i = 1 To UBound(MenuGrps)
450                       If MenuGrps(i).IncludeInToolbar Then
460                           AddGroup2Tree i, Nothing
470                       End If
480                   Next i
490                   If .Count Then .Item(1).Selected = True
500               End With
510       End Select

520       Exit Sub

ReportError:

530       MsgBox "GetElementValues has generated an error " & Err.number & ": " & vbCrLf & _
            Err.Description & " at line " & Erl, vbCritical + vbOKOnly, "Unexpected Error"

End Sub

Private Sub AddGroup2Tree(g As Integer, p As Node)

    Dim c As Integer
    Dim nItem As ListItem
    Dim tItem As Node
    Dim tcItem As Node

    With tvGroups.Nodes
        If p Is Nothing Then
            Set tItem = .Add(, , , MenuGrps(g).Caption)
        Else
            Set tItem = .Add(p, tvwChild, , MenuGrps(g).Caption)
        End If
        tItem.Bold = True
        tItem.Tag = g
        For c = 1 To UBound(MenuCmds)
            If MenuCmds(c).Parent = g Then
                Set tcItem = .Add(tItem, tvwChild, , MenuCmds(c).Caption)
                tcItem.Tag = c
                If MenuCmds(c).Actions.onmouseover.Type = atcCascade Then
                    AddGroup2Tree MenuCmds(c).Actions.onmouseover.TargetMenu, tcItem
                End If
            End If
        Next c
        tItem.Expanded = True
    End With
    
End Sub

Private Function IsGroup(itm As Node) As Boolean

    IsGroup = (itm.Bold)

End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    Unload frmFontDialog

End Sub

Private Sub lvGroups_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = 113 And Not lvGroups.SelectedItem Is Nothing Then
        lvGroups.StartLabelEdit
    End If

End Sub

Private Sub tvGroups_AfterLabelEdit(Cancel As Integer, NewString As String)

    If IsGroup(tvGroups.SelectedItem) Then
        MenuGrps(tvGroups.SelectedItem.Tag).Caption = NewString
    Else
        MenuCmds(tvGroups.SelectedItem.Tag).Caption = NewString
    End If

End Sub

Private Sub tvGroups_KeyUp(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyF2
            If Not tvGroups.SelectedItem Is Nothing Then tvGroups.StartLabelEdit
        Case vbKeyDelete
            If Not tvGroups.SelectedItem Is Nothing Then DeleteItem
    End Select

End Sub

Private Sub tvGroups_NodeClick(ByVal Node As MSComctlLib.Node)

    txtNumCmds.Text = Node.children
    If IsGroup(Node) Then
        txtURL.Text = MenuGrps(Node.Tag).Actions.onclick.url
    Else
        txtURL.Text = MenuCmds(Node.Tag).Actions.onclick.url
    End If
    cmdBrowse.Enabled = True

End Sub

Private Sub txtGroups_GotFocus()

    With txtNumCmds
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtNumCmds_GotFocus()

    With txtNumCmds
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtNumCmds_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then cmdCreateCmds_Click

End Sub

Private Sub txtURL_KeyUp(KeyCode As Integer, Shift As Integer)

    SetURL

End Sub

Private Sub ucCP_ClickCancel()

    HideColorPicker

End Sub

Private Sub ucCP_ClickOK()

    HideColorPicker
    cmdColor(SelColorBtn).BackColor = ucCP.Color

End Sub

Private Sub HideColorPicker()

    cmdColor(0).Enabled = True
    cmdColor(1).Enabled = True
    cmdColor(2).Enabled = True
    cmdColor(3).Enabled = True

    picColorPicker.Visible = False

    RestoreButtonsState

End Sub

Private Sub SaveButtonsState()

    cmdCancel.Tag = cmdCancel.Enabled:  cmdCancel.Enabled = False
    cmdNext.Tag = cmdNext.Enabled:  cmdNext.Enabled = False
    cmdBack.Tag = cmdBack.Enabled:  cmdBack.Enabled = False
    cmdPreviewInt.Tag = cmdPreviewInt.Enabled:  cmdPreviewInt.Enabled = False

End Sub

Private Sub RestoreButtonsState()

    cmdCancel.Enabled = cmdCancel.Tag
    cmdNext.Enabled = cmdNext.Tag
    cmdBack.Enabled = cmdBack.Tag
    cmdPreviewInt.Enabled = cmdPreviewInt.Tag

End Sub

Private Sub UpdateFontLabel(Index As Integer)

    On Error Resume Next

    With lblFont(Index)
        .Caption = .FontName + ", "
        .Caption = .Caption & pt2px(.FontSize)
        If .FontBold Then
            .Caption = .Caption & ", Bold"
        End If
        If .FontItalic Then
            .Caption = .Caption & ", Italic"
        End If
        If .FontUnderline Then
            .Caption = .Caption & ", Underline"
        End If
    End With

End Sub
