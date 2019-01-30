VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DHTML Batch Compiler"
   ClientHeight    =   3540
   ClientLeft      =   6210
   ClientTop       =   5520
   ClientWidth     =   6810
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6810
   Begin VB.PictureBox picObj1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3765
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   13
      Top             =   75
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picItemIcon2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3420
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   12
      Top             =   60
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
      Left            =   3120
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   60
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.TreeView tvMenus 
      Height          =   615
      Left            =   5655
      TabIndex        =   10
      Top             =   1275
      Visible         =   0   'False
      Width           =   630
      _ExtentX        =   1111
      _ExtentY        =   1085
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   135
      ScaleHeight     =   270
      ScaleWidth      =   3990
      TabIndex        =   9
      Top             =   3112
      Visible         =   0   'False
      Width           =   3990
   End
   Begin VB.PictureBox picRsc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6315
      ScaleHeight     =   270
      ScaleWidth      =   360
      TabIndex        =   8
      Top             =   1275
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComctlLib.ListView lvProjects 
      Height          =   2400
      Left            =   135
      TabIndex        =   7
      Top             =   390
      Width           =   5340
      _ExtentX        =   9419
      _ExtentY        =   4233
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      OLEDropMode     =   1
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      OLEDropMode     =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Title"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   285
      Top             =   2940
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "0"
      DialogTitle     =   "Select Project"
      Filter          =   "DHTML Menu Builder Projects|*.dmb"
      MaxFileSize     =   1024
   End
   Begin VB.CommandButton cmdLoad 
      Caption         =   "Load..."
      Height          =   345
      Left            =   5610
      TabIndex        =   6
      Top             =   1995
      Width           =   1110
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   345
      Left            =   5610
      TabIndex        =   5
      Top             =   2445
      Width           =   1110
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   5610
      TabIndex        =   4
      Top             =   3060
      Width           =   1110
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Default         =   -1  'True
      Height          =   375
      Left            =   4380
      TabIndex        =   3
      Top             =   3060
      Width           =   1110
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   345
      Left            =   5610
      TabIndex        =   2
      Top             =   825
      Width           =   1110
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add..."
      Height          =   345
      Left            =   5610
      TabIndex        =   1
      Top             =   390
      Width           =   1110
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   2070
      Top             =   30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   120
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0000
            Key             =   "mnuMenuColor|mnuContextColor"
            Object.Tag             =   "Color"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":015C
            Key             =   "mnuMenuFont|mnuContextFont"
            Object.Tag             =   "Font"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":02B8
            Key             =   "mnuMenuCursor|mnuContextCursor"
            Object.Tag             =   "Cursor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0414
            Key             =   "mnuMenuImage|mnuContextImage"
            Object.Tag             =   "Image"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0570
            Key             =   "mnuMenuSFX|mnuContextSFX"
            Object.Tag             =   "Special Effects"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":06CA
            Key             =   "btnUp"
            Object.Tag             =   "Up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0826
            Key             =   "btnDown"
            Object.Tag             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0982
            Key             =   "mnuFileNew"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F1C
            Key             =   "mnuFileOpen"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1076
            Key             =   "mnuFileSave"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1610
            Key             =   "mnuEditCopy|mnuContextCopy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BAA
            Key             =   "mnuEditPaste|mnuContextPaste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2144
            Key             =   "mnuEditUndo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26DE
            Key             =   "mnuEditRedo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C78
            Key             =   "mnuEditDelete|mnuContextDelete|mnuTBContextDelete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3212
            Key             =   "mnuMenuAddGroup|mnuTBContextAddGroup"
            Object.Tag             =   "New Group"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37AE
            Key             =   "mnuMenuAddCommand|mnuContextAddCommand"
            Object.Tag             =   "New Command"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D4A
            Key             =   "mnuMenuAddSeparator|mnuContextAddSeparator"
            Object.Tag             =   "New Separator"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42E6
            Key             =   ""
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4882
            Key             =   ""
            Object.Tag             =   "Command"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E1E
            Key             =   ""
            Object.Tag             =   "Separator"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53BA
            Key             =   "mnuToolsPreview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5956
            Key             =   "mnuToolsHotSpotsEditor"
            Object.Tag             =   "HotSpots Editor"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5EF2
            Key             =   "mnuFileProjProp"
            Object.Tag             =   "Properties"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":628C
            Key             =   "mnuToolsCompile"
            Object.Tag             =   "Compile"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":70DE
            Key             =   "mnuHelpUpgrade"
            Object.Tag             =   "Upgrade"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7686
            Key             =   "mnuHelpXFXFAQ"
            Object.Tag             =   "FAQ"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":79DA
            Key             =   "mnuHelpXFXNews"
            Object.Tag             =   "News"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7D2E
            Key             =   "mnuHelpXFXSupport"
            Object.Tag             =   "Support"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8082
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":83D6
            Key             =   "mnuHelpXFXPublicForum"
            Object.Tag             =   "Forum"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":872A
            Key             =   "mnuEditFind"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8884
            Key             =   "NoEvents"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8E1E
            Key             =   "OverCascade"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":93B8
            Key             =   "Click"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9952
            Key             =   "ClickCascade"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9EEC
            Key             =   "DoubleClick"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A486
            Key             =   "DoubleClickCascade"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA20
            Key             =   "Over"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AFBA
            Key             =   "Disabled"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B554
            Key             =   "mnuMenuMargins|mnuContextMargins"
            Object.Tag             =   "Margins"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B8EE
            Key             =   "mnuToolsPublish"
            Object.Tag             =   "Publish"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E0A0
            Key             =   "mnuMenuSound"
            Object.Tag             =   "Sounds"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E1FA
            Key             =   "GClick"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E794
            Key             =   "GOver"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ED2E
            Key             =   "GOverCascade"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F2C8
            Key             =   "GNoEvents"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F862
            Key             =   "GClickCascade"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FDFC
            Key             =   "GDoubleClick"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10396
            Key             =   "GDoubleClickCascade"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10930
            Key             =   "GDisabled"
            Object.Tag             =   "DisabledGroup"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10ECA
            Key             =   "mnuEditRename|mnuContextRename|mnuTBContextRename"
            Object.Tag             =   "Rename"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11264
            Key             =   "mnuEditPreferences"
            Object.Tag             =   "Preferences"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115FE
            Key             =   "mnuFileRF"
            Object.Tag             =   "DMBProjectIcon"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14308
            Key             =   "EmptyIcon"
            Object.Tag             =   "EmptyIcon"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":146A2
            Key             =   "mnuToolsAddInEditor"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14A3C
            Key             =   ""
            Object.Tag             =   "Left"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14B96
            Key             =   ""
            Object.Tag             =   "Right"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14CF0
            Key             =   ""
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14E4A
            Key             =   ""
            Object.Tag             =   "Over"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14FA4
            Key             =   ""
            Object.Tag             =   "Font Bold"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":150B6
            Key             =   ""
            Object.Tag             =   "Font Italic"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":151C8
            Key             =   ""
            Object.Tag             =   "Font Underline"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":152DA
            Key             =   ""
            Object.Tag             =   "Size"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153EC
            Key             =   ""
            Object.Tag             =   "Font Name"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":154FE
            Key             =   ""
            Object.Tag             =   "Toolbar Item"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15658
            Key             =   ""
            Object.Tag             =   "Target Frame"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":157B2
            Key             =   ""
            Object.Tag             =   "Leading"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1590C
            Key             =   ""
            Object.Tag             =   "Group Alignment"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15EA6
            Key             =   ""
            Object.Tag             =   "Caption Alignment"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15FB8
            Key             =   ""
            Object.Tag             =   "Events"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16112
            Key             =   ""
            Object.Tag             =   "Frame"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1626C
            Key             =   ""
            Object.Tag             =   "Line"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16806
            Key             =   ""
            Object.Tag             =   "Highlight Effects"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16DA0
            Key             =   ""
            Object.Tag             =   "Transparency"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16EFA
            Key             =   ""
            Object.Tag             =   "Shadow"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17054
            Key             =   ""
            Object.Tag             =   "All Properties"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":171AE
            Key             =   ""
            Object.Tag             =   "EventOver"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17308
            Key             =   ""
            Object.Tag             =   "EventDoubleClick"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17462
            Key             =   ""
            Object.Tag             =   "EventClick"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":175BC
            Key             =   ""
            Object.Tag             =   "URL"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17B56
            Key             =   ""
            Object.Tag             =   "Action Type"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":180F0
            Key             =   ""
            Object.Tag             =   "Border Size"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1848A
            Key             =   ""
            Object.Tag             =   "Command Horizontal Margin"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18824
            Key             =   ""
            Object.Tag             =   "Command Vertical Margin"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18BBE
            Key             =   ""
            Object.Tag             =   "Vertical Margin"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18F58
            Key             =   ""
            Object.Tag             =   "Horizontal Margin"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":192F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1968C
            Key             =   ""
            Object.Tag             =   "New Window"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19A26
            Key             =   ""
            Object.Tag             =   "Colored Borders"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19DC0
            Key             =   ""
            Object.Tag             =   "Group Width"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A15A
            Key             =   ""
            Object.Tag             =   "Group Height"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A4F4
            Key             =   ""
            Object.Tag             =   "Text Color"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AA8E
            Key             =   ""
            Object.Tag             =   "Back Color"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B028
            Key             =   ""
            Object.Tag             =   "Status Text"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B5C2
            Key             =   ""
            Object.Tag             =   "Caption"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BB5C
            Key             =   "mnuRegisterUnlock"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C836
            Key             =   ""
            Object.Tag             =   "Commands Layout"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CBD0
            Key             =   ""
            Object.Tag             =   "Group Effects"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D16A
            Key             =   ""
            Object.Tag             =   "Overlay"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D704
            Key             =   ""
            Object.Tag             =   "HS-Text"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DC9E
            Key             =   ""
            Object.Tag             =   "HS-Image"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E238
            Key             =   ""
            Object.Tag             =   "HS-DynaText"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E7D2
            Key             =   "mnuHelpContents"
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ED6C
            Key             =   "www"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F306
            Key             =   "mnuMenuToolbarProperties|mnuToolsToolbarsEditor|mnuTBContextToolbarProperties"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F460
            Key             =   "mnuToolsLCMan|mnuToolsLC"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F5BA
            Key             =   "mnuMenuAddSubGroup|mnuContextAddSubGroup"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FB54
            Key             =   "mnuMenuLength|mnuContextLength"
            Object.Tag             =   "SeparatorLength"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":200EE
            Key             =   ""
            Object.Tag             =   "Justify"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20688
            Key             =   ""
            Object.Tag             =   "Toolbar Alignment"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20C22
            Key             =   ""
            Object.Tag             =   "Spanning"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20D7C
            Key             =   ""
            Object.Tag             =   "Follow Scrolling"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20ED6
            Key             =   ""
            Object.Tag             =   "Toolbar Offset"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21030
            Key             =   "mnuMenuAddToolbar|mnuTBContextAddToolbar"
            Object.Tag             =   "Add Toolbar"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2118A
            Key             =   "mnuMenuRemoveToolbar"
            Object.Tag             =   "Remove Toolbar"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":212E4
            Key             =   "mnuEditFindReplace"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2187E
            Key             =   "mnuHelpSearch"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21E18
            Key             =   "mnuHelpXFXHomePage"
            Object.Tag             =   "xFXLogo"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":221B2
            Key             =   ""
            Object.Tag             =   "Scrolling"
         EndProperty
      EndProperty
   End
   Begin VB.Label lblProjects 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Projects"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   585
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const GW_HWNDFIRST = 0
Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Private Declare Function GetWindow Lib "user32" _
  (ByVal hwnd As Long, _
   ByVal wCmd As Long) As Long
  
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Declare Function GetWindowThreadProcessId Lib "user32" _
  (ByVal hwnd As Long, _
   lpdwProcessId As Long) As Long
    
Private Declare Function BringWindowToTop Lib "user32" _
  (ByVal hwnd As Long) As Long
    
Private Declare Function SetWindowText Lib "user32" _
   Alias "SetWindowTextA" _
  (ByVal hwnd As Long, ByVal _
   lpString As String) As Long
   
Private Declare Function GetClassName Lib "user32" _
   Alias "GetClassNameA" _
  (ByVal hwnd As Long, _
   ByVal lpClassName As String, _
   ByVal nMaxCount As Long) As Long
   
Private Declare Function SendMessage Lib "user32" _
  Alias "SendMessageA" _
  (ByVal hwnd As Long, _
   ByVal wMsg As Long, _
   ByVal wParam As Long, _
   lParam As Any) As Long
   
Private Const WM_SETTEXT = &HC
Private dmbVer As Long

Private Sub cmdAdd_Click()

    MsgBox "Not implemented" + vbCrLf + "Drag & Drop projects to add them to the batch", vbInformation

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub SetBtnsState(bState As Boolean)

    Dim ctrl As Control
    
    For Each ctrl In Controls
        If TypeOf ctrl Is CommandButton Then
            ctrl.Enabled = bState
        End If
    Next ctrl

End Sub

Private Sub cmdCompile_Click()

    Dim nItem As ListItem
    Dim dmbcPath As String
    Dim tID As Long
    Dim p As ProjectDef
    Dim fn As String
    
    SetBtnsState False
    dmbcPath = AppPath + "dmbc.exe "
    
    For Each nItem In lvProjects.ListItems
        FloodPanel.Caption = "Compiling: " + nItem.Text
        FloodPanel.Value = (nItem.Index / lvProjects.ListItems.Count) * 100
    
        nItem.Selected = True
        nItem.EnsureVisible
        
        p = GetProjectProperties(nItem.Tag)
        Close #ff
        fn = p.JSFileName + ".js"
        If p.CodeOptimization <> cocDEBUG Then fn = "ie" + fn
        fn = p.UserConfigs(p.DefaultConfig).CompiledPath + fn
        If FileExists(fn) Then Kill fn
        
        tID = Shell(dmbcPath + nItem.Tag + " /silent /compile /close", vbNormalFocus)
        
        DoEvents
        
        Me.SetFocus
        
        While GethWndFromProcessID(tID) <> 0
            DoEvents
        Wend
        If Not FileExists(fn) Then
            If MsgBox("An error has occured compiling the " + nItem.Tag + " project" + vbCrLf + vbCrLf + "Do you want to continue?", vbQuestion + vbYesNo) = vbNo Then
                Exit For
            End If
        End If
    Next nItem
    
    FloodPanel.Value = 0
    
    SetBtnsState True

End Sub

Private Sub cmdLoad_Click()
    
    MsgBox "Not implemented", vbInformation

End Sub

Private Sub cmdRemove_Click()

    MsgBox "Not implemented", vbInformation

End Sub

Private Sub cmdSave_Click()

    MsgBox "Not implemented", vbInformation

End Sub

Private Sub Form_Load()

    lvProjects.ColumnHeaders(1).Width = lvProjects.Width - (22 * 15)
    Set FloodPanel.PictureControl = picFlood
    
    AppPath = Long2Short(GetSetting("DMB", "RegInfo", "InstallPath"))
    dmbVer = GetCurProjectVersion

End Sub

Private Sub lvProjects_OLEDragDrop(Data As MSComctlLib.DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    Dim p As ProjectDef
    Dim nItem As ListItem

    ReDim Files(Data.Files.Count)
    For i = 1 To Data.Files.Count
        If FileExists(Data.Files(i)) And InStr(Data.Files(i), ":") Then
            If GetFileExtension(Data.Files(i)) = "dmb" Then
                p = GetProjectProperties(Data.Files(i))
                Close #ff
                If p.version >= dmbVer Then
                    Set nItem = lvProjects.ListItems.Add(, , IIf(p.Name = "", GetFileName(p.FileName, True), p.Name))
                    nItem.Tag = Long2Short(Data.Files(i))
                Else
                    MsgBox "The project " + p.FileName + " could not be added because its from a previous version", vbInformation + vbOKOnly
                End If
            End If
        End If
    Next i

End Sub

Private Function GethWndFromProcessID(hProcessIDToFind As Long) As Long

    Dim hWndDesktop As Long
    Dim hWndChild As Long
    Dim hWndChildProcessID As Long
    
    On Local Error GoTo GethWndFromProcessID_Error
    
   'get the handle to the desktop
    hWndDesktop = GetDesktopWindow()
    
   'get the first child under the desktop
    hWndChild = GetWindow(hWndDesktop, GW_CHILD)
    
   'hwndchild will = 0 when no more child windows are found
    Do While hWndChild <> 0
    
       'get the ThreadProcessID of the window
        Call GetWindowThreadProcessId(hWndChild, hWndChildProcessID)
        
       'if it matches the target, exit returning that value
        If hWndChildProcessID = hProcessIDToFind Then
            GethWndFromProcessID = hWndChild
            Exit Do
        End If
        
       'not found, so get the next hwnd
        hWndChild = GetWindow(hWndChild, GW_HWNDNEXT)
        
    Loop
    
Exit Function

GethWndFromProcessID_Error:
    GethWndFromProcessID = 0
    Exit Function
    
End Function

