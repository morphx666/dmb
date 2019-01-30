VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DHTML Menu Builder Project Compiler"
   ClientHeight    =   7305
   ClientLeft      =   2340
   ClientTop       =   4740
   ClientWidth     =   11010
   ControlBox      =   0   'False
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
   ScaleHeight     =   7305
   ScaleWidth      =   11010
   Begin VB.PictureBox picObj1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2460
      ScaleHeight     =   270
      ScaleWidth      =   360
      TabIndex        =   14
      Top             =   3600
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComctlLib.TreeView tvMenus 
      Height          =   6765
      Left            =   5700
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   375
      WhatsThisHelpID =   20000
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   11933
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
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   90
      TabIndex        =   13
      Top             =   3120
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Compile"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3285
      TabIndex        =   5
      Top             =   3120
      Width           =   900
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   4650
      Top             =   255
   End
   Begin VB.PictureBox picItemIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   180
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   12
      Top             =   1410
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
      Left            =   480
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   11
      Top             =   1410
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox picRsc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   2475
      ScaleHeight     =   270
      ScaleWidth      =   360
      TabIndex        =   10
      Top             =   3135
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1935
      Top             =   3030
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   825
      ScaleHeight     =   270
      ScaleWidth      =   3990
      TabIndex        =   8
      Top             =   2520
      Visible         =   0   'False
      Width           =   3990
   End
   Begin xfxLine3D.ucLine3D uc3DLine2 
      Height          =   30
      Left            =   15
      TabIndex        =   7
      Top             =   2985
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4305
      TabIndex        =   6
      Top             =   3120
      Width           =   900
   End
   Begin VB.ListBox lstConfigs 
      Height          =   1185
      Left            =   825
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1125
      Width           =   3990
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   825
      TabIndex        =   2
      Top             =   705
      Width           =   3990
      _ExtentX        =   7038
      _ExtentY        =   53
   End
   Begin VB.TextBox txtProjectName 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   810
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "Project Name"
      Top             =   390
      Width           =   3990
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   4545
      Top             =   1950
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
            Picture         =   "frmMain.frx":0CCA
            Key             =   "mnuMenuColor|mnuContextColor"
            Object.Tag             =   "Color"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0E26
            Key             =   "mnuMenuFont|mnuContextFont"
            Object.Tag             =   "Font"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F82
            Key             =   "mnuMenuCursor|mnuContextCursor"
            Object.Tag             =   "Cursor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10DE
            Key             =   "mnuMenuImage|mnuContextImage"
            Object.Tag             =   "Image"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":123A
            Key             =   "mnuMenuSFX|mnuContextSFX"
            Object.Tag             =   "Special Effects"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1394
            Key             =   "btnUp"
            Object.Tag             =   "Up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F0
            Key             =   "btnDown"
            Object.Tag             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":164C
            Key             =   "mnuFileNew"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BE6
            Key             =   "mnuFileOpen"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D40
            Key             =   "mnuFileSave"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22DA
            Key             =   "mnuEditCopy|mnuContextCopy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2874
            Key             =   "mnuEditPaste|mnuContextPaste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E0E
            Key             =   "mnuEditUndo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33A8
            Key             =   "mnuEditRedo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3942
            Key             =   "mnuEditDelete|mnuContextDelete|mnuTBContextDelete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3EDC
            Key             =   "mnuMenuAddGroup|mnuTBContextAddGroup"
            Object.Tag             =   "New Group"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4478
            Key             =   "mnuMenuAddCommand|mnuContextAddCommand"
            Object.Tag             =   "New Command"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A14
            Key             =   "mnuMenuAddSeparator|mnuContextAddSeparator"
            Object.Tag             =   "New Separator"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4FB0
            Key             =   ""
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":554C
            Key             =   ""
            Object.Tag             =   "Command"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5AE8
            Key             =   ""
            Object.Tag             =   "Separator"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6084
            Key             =   "mnuToolsPreview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6620
            Key             =   "mnuToolsHotSpotsEditor"
            Object.Tag             =   "HotSpots Editor"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6BBC
            Key             =   "mnuFileProjProp"
            Object.Tag             =   "Properties"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6F56
            Key             =   "mnuToolsCompile"
            Object.Tag             =   "Compile"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7DA8
            Key             =   "mnuHelpUpgrade"
            Object.Tag             =   "Upgrade"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8350
            Key             =   "mnuHelpXFXFAQ"
            Object.Tag             =   "FAQ"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":86A4
            Key             =   "mnuHelpXFXNews"
            Object.Tag             =   "News"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":89F8
            Key             =   "mnuHelpXFXSupport"
            Object.Tag             =   "Support"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8D4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":90A0
            Key             =   "mnuHelpXFXPublicForum"
            Object.Tag             =   "Forum"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":93F4
            Key             =   "mnuEditFind"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":954E
            Key             =   "NoEvents"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9AE8
            Key             =   "OverCascade"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A082
            Key             =   "Click"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A61C
            Key             =   "ClickCascade"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ABB6
            Key             =   "DoubleClick"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B150
            Key             =   "DoubleClickCascade"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B6EA
            Key             =   "Over"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BC84
            Key             =   "Disabled"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C21E
            Key             =   "mnuMenuMargins|mnuContextMargins"
            Object.Tag             =   "Margins"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C5B8
            Key             =   "mnuToolsPublish"
            Object.Tag             =   "Publish"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":ED6A
            Key             =   "mnuMenuSound"
            Object.Tag             =   "Sounds"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":EEC4
            Key             =   "GClick"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F45E
            Key             =   "GOver"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":F9F8
            Key             =   "GOverCascade"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":FF92
            Key             =   "GNoEvents"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1052C
            Key             =   "GClickCascade"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10AC6
            Key             =   "GDoubleClick"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11060
            Key             =   "GDoubleClickCascade"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":115FA
            Key             =   "GDisabled"
            Object.Tag             =   "DisabledGroup"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11B94
            Key             =   "mnuEditRename|mnuContextRename|mnuTBContextRename"
            Object.Tag             =   "Rename"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11F2E
            Key             =   "mnuEditPreferences"
            Object.Tag             =   "Preferences"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":122C8
            Key             =   "mnuFileRF"
            Object.Tag             =   "DMBProjectIcon"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14FD2
            Key             =   "EmptyIcon"
            Object.Tag             =   "EmptyIcon"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1536C
            Key             =   "mnuToolsAddInEditor"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15706
            Key             =   ""
            Object.Tag             =   "Left"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15860
            Key             =   ""
            Object.Tag             =   "Right"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":159BA
            Key             =   ""
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15B14
            Key             =   ""
            Object.Tag             =   "Over"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15C6E
            Key             =   ""
            Object.Tag             =   "Font Bold"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15D80
            Key             =   ""
            Object.Tag             =   "Font Italic"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15E92
            Key             =   ""
            Object.Tag             =   "Font Underline"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15FA4
            Key             =   ""
            Object.Tag             =   "Size"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":160B6
            Key             =   ""
            Object.Tag             =   "Font Name"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":161C8
            Key             =   ""
            Object.Tag             =   "Toolbar Item"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16322
            Key             =   ""
            Object.Tag             =   "Target Frame"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1647C
            Key             =   ""
            Object.Tag             =   "Leading"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":165D6
            Key             =   ""
            Object.Tag             =   "Group Alignment"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16B70
            Key             =   ""
            Object.Tag             =   "Caption Alignment"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16C82
            Key             =   ""
            Object.Tag             =   "Events"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16DDC
            Key             =   ""
            Object.Tag             =   "Frame"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":16F36
            Key             =   ""
            Object.Tag             =   "Line"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":174D0
            Key             =   ""
            Object.Tag             =   "Highlight Effects"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17A6A
            Key             =   ""
            Object.Tag             =   "Transparency"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17BC4
            Key             =   ""
            Object.Tag             =   "Shadow"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17D1E
            Key             =   ""
            Object.Tag             =   "All Properties"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17E78
            Key             =   ""
            Object.Tag             =   "EventOver"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17FD2
            Key             =   ""
            Object.Tag             =   "EventDoubleClick"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1812C
            Key             =   ""
            Object.Tag             =   "EventClick"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18286
            Key             =   ""
            Object.Tag             =   "URL"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18820
            Key             =   ""
            Object.Tag             =   "Action Type"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18DBA
            Key             =   ""
            Object.Tag             =   "Border Size"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19154
            Key             =   ""
            Object.Tag             =   "Command Horizontal Margin"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":194EE
            Key             =   ""
            Object.Tag             =   "Command Vertical Margin"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19888
            Key             =   ""
            Object.Tag             =   "Vertical Margin"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19C22
            Key             =   ""
            Object.Tag             =   "Horizontal Margin"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A356
            Key             =   ""
            Object.Tag             =   "New Window"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A6F0
            Key             =   ""
            Object.Tag             =   "Colored Borders"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AA8A
            Key             =   ""
            Object.Tag             =   "Group Width"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AE24
            Key             =   ""
            Object.Tag             =   "Group Height"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B1BE
            Key             =   ""
            Object.Tag             =   "Text Color"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B758
            Key             =   ""
            Object.Tag             =   "Back Color"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BCF2
            Key             =   ""
            Object.Tag             =   "Status Text"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C28C
            Key             =   ""
            Object.Tag             =   "Caption"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C826
            Key             =   "mnuRegisterUnlock"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D500
            Key             =   ""
            Object.Tag             =   "Commands Layout"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D89A
            Key             =   ""
            Object.Tag             =   "Group Effects"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DE34
            Key             =   ""
            Object.Tag             =   "Overlay"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E3CE
            Key             =   ""
            Object.Tag             =   "HS-Text"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E968
            Key             =   ""
            Object.Tag             =   "HS-Image"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1EF02
            Key             =   ""
            Object.Tag             =   "HS-DynaText"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F49C
            Key             =   "mnuHelpContents"
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FA36
            Key             =   "www"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FFD0
            Key             =   "mnuMenuToolbarProperties|mnuToolsToolbarsEditor|mnuTBContextToolbarProperties"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2012A
            Key             =   "mnuToolsLCMan|mnuToolsLC"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20284
            Key             =   "mnuMenuAddSubGroup|mnuContextAddSubGroup"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2081E
            Key             =   "mnuMenuLength|mnuContextLength"
            Object.Tag             =   "SeparatorLength"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20DB8
            Key             =   ""
            Object.Tag             =   "Justify"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21352
            Key             =   ""
            Object.Tag             =   "Toolbar Alignment"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":218EC
            Key             =   ""
            Object.Tag             =   "Spanning"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21A46
            Key             =   ""
            Object.Tag             =   "Follow Scrolling"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21BA0
            Key             =   ""
            Object.Tag             =   "Toolbar Offset"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21CFA
            Key             =   "mnuMenuAddToolbar|mnuTBContextAddToolbar"
            Object.Tag             =   "Add Toolbar"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21E54
            Key             =   "mnuMenuRemoveToolbar"
            Object.Tag             =   "Remove Toolbar"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21FAE
            Key             =   "mnuEditFindReplace"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22548
            Key             =   "mnuHelpSearch"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22AE2
            Key             =   "mnuHelpXFXHomePage"
            Object.Tag             =   "xFXLogo"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22E7C
            Key             =   ""
            Object.Tag             =   "Scrolling"
         EndProperty
      EndProperty
   End
   Begin VB.Image imgLogo 
      Height          =   480
      Left            =   135
      Picture         =   "frmMain.frx":23216
      Top             =   450
      Width           =   480
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Available Configurations"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   825
      TabIndex        =   3
      Top             =   855
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Project Name"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   825
      TabIndex        =   0
      Top             =   135
      Width           =   1110
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsUpdating As Boolean
Dim FileName As String

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdHelp_Click()

    ShowHelp

End Sub

Private Sub cmdOK_Click()

          Dim i As Integer
          
10        On Error Resume Next
          
20        Err.Clear
          
30        Screen.MousePointer = vbHourglass
40        Me.Enabled = False
          
50        If LoadMenu(FileName) Then
60            If Project.FileName = "" Then Project.FileName = FileName
70            For i = 1 To UBound(Project.Toolbars)
80                Project.Toolbars(i).Compile = True
90            Next i
100           For i = 1 To UBound(MenuCmds)
110               MenuCmds(i).Compile = True
120           Next i
130           For i = 1 To UBound(MenuGrps)
140               MenuGrps(i).Compile = True
150           Next i
                
160           CompileProject MenuGrps, MenuCmds, Project, Preferences, params, False
170           With Project
180               .DOMCode = ""
190               .DOMFramesCode = ""
200               .NSCode = ""
210               .NSFramesCode = ""
220           End With
230       End If
          
240       If Err.number > 0 And Erl <> 0 Then
250           MsgBox "Errors have occured while compiling the project:" + vbCrLf + "Error " & Err.number & ": " & Err.Description & " at " & Erl, vbCritical + vbOKOnly, "Compilation Errors"
260       End If
          
270       Screen.MousePointer = vbDefault
280       Me.Enabled = True

End Sub

Private Function LoadMenu(Optional File As String) As Boolean

    Dim sStr As String
    Dim Ans As Integer
    Dim nLines As Integer
    Dim cLine As Integer
        
    On Error GoTo chkError
    
    ff = FreeFile
    Open File For Input As ff
        Do Until EOF(ff) Or sStr = "[RSC]"
            Line Input #ff, sStr
            nLines = nLines + 1
        Loop
    Close ff
    nLines = nLines - 2
    FloodPanel.Caption = "Loading"
    
    Project = GetProjectProperties(File)
    
    Erase MenuGrps: ReDim MenuGrps(0)
    Erase MenuCmds: ReDim MenuCmds(0)
    
'    If EOF(ff) Then GoTo ExitSub
'    Line Input #ff, sStr
'    Do Until EOF(ff) Or sStr = "[RSC]"
'        AddMenuGroup Mid$(sStr, 4)
'        Do Until EOF(ff) Or sStr = "[RSC]"
'            Line Input #ff, sStr
'            cLine = cLine + 1: FloodPanel.Value = (cLine / nLines) * 100
'            If Left$(sStr, 3) = "[C]" Then
'                AddMenuCommand Mid$(sStr, 6), True
'            Else
'                Exit Do
'            End If
'        Loop
'        If EOF(ff) Then Exit Do
'    Loop
    
    If (LOF(ff) = Loc(ff)) Then GoTo ExitSub
    Line Input #ff, sStr
    Do Until (LOF(ff) = Loc(ff)) Or sStr = "[RSC]"
        If sStr <> "" Then AddMenuGroup Mid$(sStr, 4)
        Do Until (LOF(ff) = Loc(ff)) Or sStr = "[RSC]"
            Line Input #ff, sStr
            If sStr <> "" Then
                If Left$(sStr, 3) = "[C]" Then
                    AddMenuCommand Mid$(sStr, 6), True
                Else
                    Exit Do
                End If
            End If
        Loop
        If (LOF(ff) = Loc(ff)) Then Exit Do
    Loop
    
    LoadMenu = True
ExitSub:
    Close ff
    
    FloodPanel.Value = 0
   
    Exit Function
    
chkError:
    If Err.number = 53 Or Err.number = 76 Then
        MsgBox "The project could not be opened because the file " + Project.FileName + " does not exists", vbCritical + vbOKOnly, "Error Opening Project"
    Else
        MsgBox "The project '" + Project.FileName + "' could not be opened. Error (" & Err.number & ") " + Err.Description, vbCritical + vbOKOnly, "Error Opening Project"
    End If
    Project.HasChanged = False
    GoTo ExitSub

End Function

Private Sub GetPrgPrefs()
    
    USER = GetSetting("DMB", "RegInfo", "User", "DEMO")
    COMPANY = GetSetting("DMB", "RegInfo", "Company", "DEMO")
    USERSN = GetSetting("DMB", "RegInfo", "SerialNumber", "")
    
    If USER = "DEMO" Or USER = "" Or USERSN = "" Then End
    
    With Preferences
        .AutoRecover = GetSetting("DMB", "Preferences", "AutoRecover", True)
        .OpenLastProject = GetSetting("DMB", "Preferences", "OpenLastProject", True)
        .SepHeight = GetSetting("DMB", "Preferences", "SepHeight", 13)
        .ShowNag = GetSetting("DMB", "Preferences", "ShowNag", True)
        .ShowWarningAddInEditor = GetSetting("DMB", "Preferences", "ShowWarningAIE", True)
        .CommandsInheritance = GetSetting("DMB", "Preferences", "CmdInh", icFirst)
        .GroupsInheritance = GetSetting("DMB", "Preferences", "GrpInh", icFirst)
        .UseLivePreview = GetSetting("DMB", "Preferences", "UseLivePreview", True)
        .EnableUndoRedo = GetSetting("DMB", "Preferences", "DisableUR", True)
        .ImgSpace = Val(GetSetting("DMB", "Preferences", "ImgSpace", 4))
    End With
    
End Sub

Private Sub Form_Load()

    Width = 5430
    Height = 3975

    AppPath = AddTrailingSlash(App.Path, "\")
    Preferences.Language = "eng"

    LoadLocalizedStrings
    
    DoEvents
    
    CenterForm Me
    
    Me.Visible = False
    
    tmrInit.Enabled = True

End Sub

Private Sub lstConfigs_Click()

    If IsUpdating Then Exit Sub
    lstConfigs_ItemCheck lstConfigs.ListIndex
    
    UpdateItemsLinks

End Sub

Private Sub lstConfigs_ItemCheck(Item As Integer)

    Dim i As Integer
    
    If IsUpdating Then Exit Sub
    IsUpdating = True
    
    For i = 0 To lstConfigs.ListCount - 1
        lstConfigs.Selected(i) = False
    Next i
    lstConfigs.Selected(Item) = True
    
    IsUpdating = False

End Sub

Private Sub tmrClose_Timer()

    Unload Me

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

Private Sub MkDir2(ByVal d As String)

    Dim i As Integer
    Dim sd() As String
    Dim ss As String
    
    sd = Split(d, "\")
    For i = 0 To UBound(sd)
        ss = Replace(ss + sd(i) + "\", "\\", "\")
        If i > 0 Then
            If Not FolderExists(ss) Then MkDir ss
        End If
    Next i

End Sub

Private Sub tmrInit_Timer()

    Dim i As Integer
    
    tmrInit.Enabled = False
    
    GetPrgPrefs
    
    cSep = Chr(255) + Chr(255)
    
    AppPath = GetSetting("DMB", "RegInfo", "InstallPath")
    SetupTempFolders
    SetTemplateDefaults
    
    Set FloodPanel.PictureControl = picFlood
    Dim UIObjects(1 To 3) As Object
    Set UIObjects(1) = frmMain
    Set UIObjects(2) = FloodPanel
    Set UIObjects(3) = frmMain
    SetUI UIObjects
    
    Dim VarObjects(1 To 7) As Variant
    VarObjects(1) = AppPath
    VarObjects(2) = ""
    VarObjects(3) = ""
    VarObjects(4) = GetTEMPPath
    VarObjects(5) = cSep
    VarObjects(6) = nwdPar
    VarObjects(7) = StatesPath
    SetVars VarObjects
    
    Caption = "DHTML Menu Builder Project Compiler " + GetFileVersion(AppPath + "dmbc.exe")

    If InStr(Command, "/") Then
        FileName = CStr(Left(Command, InStr(Command, "/") - 2))
    Else
        FileName = CStr(Command)
    End If
    If FileName <> "" Then
        Project = GetProjectProperties(FileName)
        Close ff
        
        txtProjectName.Text = Project.Name
        
        For i = 0 To UBound(Project.UserConfigs)
            lstConfigs.AddItem Project.UserConfigs(i).Name
            lstConfigs.Selected(lstConfigs.NewIndex) = (Project.DefaultConfig = i)
        Next i
        
        ProcessCommandLine Mid(Command$, Len(FileName) + 2)
    Else
        ShowHelp
        tmrClose.Enabled = True
    End If

End Sub

Private Sub ProcessCommandLine(cLine As String)

    Dim op() As String
    Dim i As Integer
    Dim k As Integer
    Dim p() As String
    Dim IsSilent As Boolean
    Dim t As Integer

    op = Split(cLine, "/")
    
    DoEvents
    
    t = UBound(op)
    If t = -1 Then
        Me.Visible = True
    Else
        For i = 1 To t
            If InStr(op(i), ":") Then
                p() = Split(op(i), ":")
                p(1) = Trim(p(1))
                Select Case p(0)
                    Case "config"
                        For k = 0 To lstConfigs.ListCount - 1
                            If LCase(p(1)) = LCase(lstConfigs.List(k)) Then
                                lstConfigs_ItemCheck k
                                Exit For
                            End If
                        Next k
                End Select
            Else
                Select Case Trim(LCase(op(i)))
                    Case "compile"
                        Me.Visible = Not IsSilent
                        cmdOK_Click
                    Case "close"
                        tmrClose.Enabled = True
                    Case "silent"
                        IsSilent = True
                    Case "help"
                        ShowHelp
                End Select
            End If
        Next i
    End If
    
End Sub

Private Sub ShowHelp()

    MsgBox "dmbc.exe FileName [/silent] [/config:Name] [/compile] [/close] [/help]" + vbCrLf + vbCrLf + _
            "/silent" + vbTab + vbTab + "-> Hides the Project Compiler window. When using this option, make sure you include the /close argument" + vbCrLf + _
            "/config:Name" + vbTab + "-> Sets the default configuration" + vbCrLf + _
            "/compile" + vbTab + vbTab + "-> Compiles the loaded project specified by FileName" + vbCrLf + _
            "/close" + vbTab + vbTab + "-> Closes the Project Compiler after processing all the command line arguments" + vbCrLf + _
            "/help" + vbTab + vbTab + "-> Displays this dialog" + vbCrLf + vbCrLf + _
            "NOTE: The command line arguments are processed in the same order as entered" + vbCrLf + vbCrLf + _
            "The Project Compiler can also be accessed from Explorer by rightclicking a" + vbCrLf + "DHTML Menu Builder project and selecting 'Compile Project' from the Context Menu", _
            vbInformation + vbOKOnly, "Project Compiler " + GetFileVersion(AppPath + "dmbc.exe") + " - Command Line Arguments"

End Sub
