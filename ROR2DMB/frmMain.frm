VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "ROR2DMB"
   ClientHeight    =   6855
   ClientLeft      =   2835
   ClientTop       =   4140
   ClientWidth     =   10680
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
   ScaleHeight     =   6855
   ScaleWidth      =   10680
   Begin VB.ComboBox cmbRoot 
      Height          =   315
      Left            =   1845
      Style           =   2  'Dropdown List
      TabIndex        =   16
      Top             =   4635
      Width           =   2595
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Enabled         =   0   'False
      Height          =   405
      Left            =   8445
      TabIndex        =   11
      Top             =   4935
      Width           =   1020
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   405
      Left            =   9555
      TabIndex        =   10
      Top             =   4935
      Width           =   1020
   End
   Begin VB.PictureBox picDiv 
      Height          =   1020
      Left            =   5415
      MousePointer    =   9  'Size W E
      ScaleHeight     =   960
      ScaleWidth      =   60
      TabIndex        =   9
      Top             =   4245
      Width           =   120
   End
   Begin VB.PictureBox picObj1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   240
      ScaleHeight     =   270
      ScaleWidth      =   360
      TabIndex        =   8
      Top             =   5460
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picItemIcon2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   1110
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   7
      Top             =   4995
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
      Left            =   810
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   6
      Top             =   4995
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
      Left            =   180
      ScaleHeight     =   270
      ScaleWidth      =   360
      TabIndex        =   5
      Top             =   4995
      Visible         =   0   'False
      Width           =   360
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   2985
      Top             =   75
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
            Picture         =   "frmMain.frx":2CFA
            Key             =   "mnuMenuColor|mnuContextColor"
            Object.Tag             =   "Color"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2E56
            Key             =   "mnuMenuFont|mnuContextFont"
            Object.Tag             =   "Font"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2FB2
            Key             =   "mnuMenuCursor|mnuContextCursor"
            Object.Tag             =   "Cursor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":310E
            Key             =   "mnuMenuImage|mnuContextImage"
            Object.Tag             =   "Image"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":326A
            Key             =   "mnuMenuSFX|mnuContextSFX"
            Object.Tag             =   "Special Effects"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":33C4
            Key             =   "btnUp"
            Object.Tag             =   "Up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3520
            Key             =   "btnDown"
            Object.Tag             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":367C
            Key             =   "mnuFileNew"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3C16
            Key             =   "mnuFileOpen"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D70
            Key             =   "mnuFileSave"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":430A
            Key             =   "mnuEditCopy|mnuContextCopy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":48A4
            Key             =   "mnuEditPaste|mnuContextPaste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4E3E
            Key             =   "mnuEditUndo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":53D8
            Key             =   "mnuEditRedo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5972
            Key             =   "mnuEditDelete|mnuContextDelete|mnuTBContextDelete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":5F0C
            Key             =   "mnuMenuAddGroup|mnuTBContextAddGroup"
            Object.Tag             =   "New Group"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":64A8
            Key             =   "mnuMenuAddCommand|mnuContextAddCommand"
            Object.Tag             =   "New Command"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6A44
            Key             =   "mnuMenuAddSeparator|mnuContextAddSeparator"
            Object.Tag             =   "New Separator"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":6FE0
            Key             =   ""
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":757C
            Key             =   ""
            Object.Tag             =   "Command"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":7B18
            Key             =   ""
            Object.Tag             =   "Separator"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":80B4
            Key             =   "mnuToolsPreview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8650
            Key             =   "mnuToolsHotSpotsEditor"
            Object.Tag             =   "HotSpots Editor"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8BEC
            Key             =   "mnuFileProjProp"
            Object.Tag             =   "Properties"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":8F86
            Key             =   "mnuToolsCompile"
            Object.Tag             =   "Compile"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":9DD8
            Key             =   "mnuHelpUpgrade"
            Object.Tag             =   "Upgrade"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A380
            Key             =   "mnuHelpXFXFAQ"
            Object.Tag             =   "FAQ"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":A6D4
            Key             =   "mnuHelpXFXNews"
            Object.Tag             =   "News"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AA28
            Key             =   "mnuHelpXFXSupport"
            Object.Tag             =   "Support"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":AD7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B0D0
            Key             =   "mnuHelpXFXPublicForum"
            Object.Tag             =   "Forum"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B424
            Key             =   "mnuEditFind"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":B57E
            Key             =   "NoEvents"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":BB18
            Key             =   "OverCascade"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C0B2
            Key             =   "Click"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":C64C
            Key             =   "ClickCascade"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":CBE6
            Key             =   "DoubleClick"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D180
            Key             =   "DoubleClickCascade"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":D71A
            Key             =   "Over"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":DCB4
            Key             =   "Disabled"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E24E
            Key             =   "mnuMenuMargins|mnuContextMargins"
            Object.Tag             =   "Margins"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":E5E8
            Key             =   "mnuToolsPublish"
            Object.Tag             =   "Publish"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10D9A
            Key             =   "mnuMenuSound"
            Object.Tag             =   "Sounds"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":10EF4
            Key             =   "GClick"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1148E
            Key             =   "GOver"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11A28
            Key             =   "GOverCascade"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":11FC2
            Key             =   "GNoEvents"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1255C
            Key             =   "GClickCascade"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":12AF6
            Key             =   "GDoubleClick"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13090
            Key             =   "GDoubleClickCascade"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1362A
            Key             =   "GDisabled"
            Object.Tag             =   "DisabledGroup"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13BC4
            Key             =   "mnuEditRename|mnuContextRename|mnuTBContextRename"
            Object.Tag             =   "Rename"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":13F5E
            Key             =   "mnuEditPreferences"
            Object.Tag             =   "Preferences"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":142F8
            Key             =   "mnuFileRF"
            Object.Tag             =   "DMBProjectIcon"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17002
            Key             =   "EmptyIcon"
            Object.Tag             =   "EmptyIcon"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1739C
            Key             =   "mnuToolsAddInEditor"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17736
            Key             =   ""
            Object.Tag             =   "Left"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17890
            Key             =   ""
            Object.Tag             =   "Right"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":179EA
            Key             =   ""
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17B44
            Key             =   ""
            Object.Tag             =   "Over"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17C9E
            Key             =   ""
            Object.Tag             =   "Font Bold"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17DB0
            Key             =   ""
            Object.Tag             =   "Font Italic"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17EC2
            Key             =   ""
            Object.Tag             =   "Font Underline"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":17FD4
            Key             =   ""
            Object.Tag             =   "Size"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":180E6
            Key             =   ""
            Object.Tag             =   "Font Name"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":181F8
            Key             =   ""
            Object.Tag             =   "Toolbar Item"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18352
            Key             =   ""
            Object.Tag             =   "Target Frame"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":184AC
            Key             =   ""
            Object.Tag             =   "Leading"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18606
            Key             =   ""
            Object.Tag             =   "Group Alignment"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18BA0
            Key             =   ""
            Object.Tag             =   "Caption Alignment"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18CB2
            Key             =   ""
            Object.Tag             =   "Events"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18E0C
            Key             =   ""
            Object.Tag             =   "Frame"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":18F66
            Key             =   ""
            Object.Tag             =   "Line"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19500
            Key             =   ""
            Object.Tag             =   "Highlight Effects"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19A9A
            Key             =   ""
            Object.Tag             =   "Transparency"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19BF4
            Key             =   ""
            Object.Tag             =   "Shadow"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19D4E
            Key             =   ""
            Object.Tag             =   "All Properties"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":19EA8
            Key             =   ""
            Object.Tag             =   "EventOver"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A002
            Key             =   ""
            Object.Tag             =   "EventDoubleClick"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A15C
            Key             =   ""
            Object.Tag             =   "EventClick"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A2B6
            Key             =   ""
            Object.Tag             =   "URL"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1A850
            Key             =   ""
            Object.Tag             =   "Action Type"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1ADEA
            Key             =   ""
            Object.Tag             =   "Border Size"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B184
            Key             =   ""
            Object.Tag             =   "Command Horizontal Margin"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B51E
            Key             =   ""
            Object.Tag             =   "Command Vertical Margin"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1B8B8
            Key             =   ""
            Object.Tag             =   "Vertical Margin"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BC52
            Key             =   ""
            Object.Tag             =   "Horizontal Margin"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1BFEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C386
            Key             =   ""
            Object.Tag             =   "New Window"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C720
            Key             =   ""
            Object.Tag             =   "Colored Borders"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CABA
            Key             =   ""
            Object.Tag             =   "Group Width"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1CE54
            Key             =   ""
            Object.Tag             =   "Group Height"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D1EE
            Key             =   ""
            Object.Tag             =   "Text Color"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1D788
            Key             =   ""
            Object.Tag             =   "Back Color"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DD22
            Key             =   ""
            Object.Tag             =   "Status Text"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E2BC
            Key             =   ""
            Object.Tag             =   "Caption"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1E856
            Key             =   "mnuRegisterUnlock"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F530
            Key             =   ""
            Object.Tag             =   "Commands Layout"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1F8CA
            Key             =   ""
            Object.Tag             =   "Group Effects"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1FE64
            Key             =   ""
            Object.Tag             =   "Overlay"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":203FE
            Key             =   ""
            Object.Tag             =   "HS-Text"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20998
            Key             =   ""
            Object.Tag             =   "HS-Image"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":20F32
            Key             =   ""
            Object.Tag             =   "HS-DynaText"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":214CC
            Key             =   "mnuHelpContents"
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":21A66
            Key             =   "www"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22000
            Key             =   "mnuMenuToolbarProperties|mnuToolsToolbarsEditor|mnuTBContextToolbarProperties"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2215A
            Key             =   "mnuToolsLCMan|mnuToolsLC"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":222B4
            Key             =   "mnuMenuAddSubGroup|mnuContextAddSubGroup"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2284E
            Key             =   "mnuMenuLength|mnuContextLength"
            Object.Tag             =   "SeparatorLength"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":22DE8
            Key             =   ""
            Object.Tag             =   "Justify"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23382
            Key             =   ""
            Object.Tag             =   "Toolbar Alignment"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2391C
            Key             =   ""
            Object.Tag             =   "Spanning"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23A76
            Key             =   ""
            Object.Tag             =   "Follow Scrolling"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23BD0
            Key             =   ""
            Object.Tag             =   "Toolbar Offset"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23D2A
            Key             =   "mnuMenuAddToolbar|mnuTBContextAddToolbar"
            Object.Tag             =   "Add Toolbar"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23E84
            Key             =   "mnuMenuRemoveToolbar"
            Object.Tag             =   "Remove Toolbar"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":23FDE
            Key             =   "mnuEditFindReplace"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24578
            Key             =   "mnuHelpSearch"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24B12
            Key             =   "mnuHelpXFXHomePage"
            Object.Tag             =   "xFXLogo"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":24EAC
            Key             =   ""
            Object.Tag             =   "Scrolling"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   495
      ScaleHeight     =   270
      ScaleWidth      =   5220
      TabIndex        =   4
      Top             =   5910
      Visible         =   0   'False
      Width           =   5220
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   4500
      Top             =   45
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   3780
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   90
      TabIndex        =   1
      Top             =   240
      Width           =   4890
   End
   Begin SmartButtonProject.SmartButton cmdBrowse 
      Height          =   315
      Left            =   5070
      TabIndex        =   2
      Top             =   225
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmMain.frx":25246
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
   Begin MSComctlLib.TreeView tvMenus 
      Height          =   1320
      Left            =   825
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5355
      Visible         =   0   'False
      WhatsThisHelpID =   20000
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2328
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
   Begin SHDocVwCtl.WebBrowser wbPreview 
      Height          =   4575
      Left            =   5520
      TabIndex        =   12
      Top             =   240
      Width           =   5115
      ExtentX         =   9022
      ExtentY         =   8070
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
   Begin MSComctlLib.TreeView tvRORMenus 
      Height          =   3060
      Left            =   90
      TabIndex        =   13
      Top             =   945
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   5398
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Label lblOpenInfo 
      BackStyle       =   0  'Transparent
      Caption         =   "Use the open icon to locate a ROR file or type the fully qualified URL to a ROR file on the Internet"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000011&
      Height          =   330
      Left            =   90
      TabIndex        =   17
      Top             =   570
      Width           =   4890
   End
   Begin VB.Label lblRootTB 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create Toolbar From"
      Height          =   195
      Left            =   105
      TabIndex        =   15
      Top             =   4590
      Width           =   1485
   End
   Begin VB.Label lblPreview 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Preview"
      Height          =   195
      Left            =   5520
      TabIndex        =   14
      Top             =   15
      Width           =   570
   End
   Begin VB.Label lblRORDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROR Document"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   15
      Width           =   1095
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileOpen 
         Caption         =   "Open..."
      End
      Begin VB.Menu mnuFileSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "Save..."
      End
      Begin VB.Menu mnuFileSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuPreview 
      Caption         =   "Menus"
      Begin VB.Menu mnuPreviewTBS 
         Caption         =   "Toolbar Style"
         Begin VB.Menu mnuPreviewTBSHorizontal 
            Caption         =   "Horizontal"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuPreviewTBSVertical 
            Caption         =   "Vertical"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim FileName As String
Dim IsDragging As Boolean
Dim mX As Integer

Dim cNode As Long
Dim tNodes As Long
Dim cGrp As Integer
Dim tGrps As Integer

Private Sub cmbRoot_Click()

    SetCtrlsState False
    
    cGrp = 0
    tGrps = 0
    CreateDMBProject
    
    SetCtrlsState True

End Sub

Private Sub cmdBrowse_Click()

    OpenROR

End Sub

Private Sub OpenROR()

    With cDlg
        .DialogTitle = "Select ROR File"
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNNoReadOnlyReturn
        .Filter = "XML ROR Files (*.xml)|*.xml"
        Err.Clear
        On Error Resume Next
        .ShowOpen
        If Err.number = 0 Then
            On Error GoTo 0
            If LenB(.FileName) <> 0 Then
                txtFileName.Text = .FileName
                ParseROR .FileName
            End If
        End If
    End With

End Sub

Private Sub ParseROR(mFileName As String)

    Dim oXMLDOM As MSXML2.DOMDocument
    
    On Error GoTo chkError
    
    If mFileName = "" Then Exit Sub
    
    wbPreview.Navigate "about:blank"
    FileName = mFileName
    
    Set oXMLDOM = New MSXML2.DOMDocument
    oXMLDOM.async = False
    oXMLDOM.Load FileName
    
    tvRORMenus.Nodes.Clear
    cmbRoot.Clear
    cNode = 0
    tNodes = 0
    
    StartParsing oXMLDOM
    AddRoots
    
ExitSub:

    FloodPanel.Value = 0
    SetCtrlsState True
    
    Exit Sub
    
chkError:
    MsgBox "Error loading/parsing: " + FileName + vbCrLf + _
            "Error " & Err.number & ": " & Err.Description, vbCritical + vbOKOnly, "Parser Error"

    GoTo ExitSub
End Sub

Private Function StartParsing(n As IXMLDOMNode) As IXMLDOMNode

    Dim sn As IXMLDOMNode
    Dim ssn As IXMLDOMNode
    Dim pNode As Node
    Dim sectStr As String
    
    On Error GoTo chkError
    
    For Each sn In n.childNodes
        If LCase(sn.baseName) = "rdf" Then
            sectStr = "resource"
            
            If LCase(sn.baseName) = "rdf" Then
                For Each ssn In sn.childNodes
                    If LCase(ssn.nodeName) = "resource" Then
                        Set pNode = tvRORMenus.Nodes.Add(, , , FileName)
                        pNode.Expanded = True
                        
                        parseNode ssn.parentNode, pNode
                        Exit For
                    End If
                Next ssn
            End If
            
            Exit Function
        End If
    Next sn
    
    If sn Is Nothing Then
        For Each sn In n.childNodes
            If LCase(sn.baseName) = "rss" Then
                Set pNode = tvRORMenus.Nodes.Add(, , , FileName)
                pNode.Expanded = True
                
                For Each ssn In sn.childNodes
                    If LCase(ssn.nodeName) = "channel" Then
                        Set pNode = tvRORMenus.Nodes.Add(pNode, tvwChild, , "Channel")
                        pNode.Expanded = True
                        pNode.Bold = True
                        
                        Set pNode = tvRORMenus.Nodes.Add(pNode, tvwChild, , getTitle(ssn, False))
                        pNode.Expanded = True
                        pNode.Bold = True
                        
                        parseNode ssn, pNode
                        Exit For
                    End If
                Next ssn
                
                Exit Function
            End If
        Next sn
    End If
    
    If sn Is Nothing Then
        Err.Raise "1001", "StartParsing", "rdf namespace not found. The ROR file appears to be damaged"
        GoTo ExitSub
    End If
    
    
    
ExitSub:
    FloodPanel.Value = 0
    
    Exit Function
    
chkError:
    MsgBox "Error loading/parsing: " + FileName + vbCrLf + _
            "Error " & Err.number & ": " & Err.Description, vbCritical + vbOKOnly, "Parser Error"

    GoTo ExitSub

End Function

Private Sub SetCtrlsState(State As Boolean)

    txtFileName.Enabled = State
    cmdBrowse.Enabled = State
    cmbRoot.Enabled = State
    cmdClose.Enabled = State
    cmdSave.Enabled = State

End Sub

Private Sub AddRoots()

    Dim n As Node
    
    For Each n In tvRORMenus.Nodes
        If n.Bold Then
            If HasSubMenus(n) Then cmbRoot.AddItem n.Text
        End If
    Next n
    
    If cmbRoot.ListCount > 0 Then
        cmbRoot.ListIndex = 0
    Else
        cmbRoot.Enabled = False
    End If

End Sub

Private Sub CreateDMBProject()
    
    ReDim MenuGrps(0)
    ReDim MenuCmds(0)
    
    tvMenus.Nodes.Clear
    
    With Project
        .Name = tvRORMenus.Nodes(1).Child.Text
        .AbsPath = ""
        .FileName = ""
        .AddIn.Name = ""
        .CodeOptimization = cocAggressive
        .RemoveImageAutoPosCode = True
        .FX = 0
        .DXFilter = "progid:DXImageTransform.Microsoft.Fade(overlap=1,duration=0.3)"
                
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
            .BackColor = vbBlack
            .Border = 0
            .BorderColor = &H0
            
            .CustX = 0
            .CustY = 0
            
            .FollowHScroll = False
            .FollowVScroll = False

            .ContentsMarginH = 0
            .ContentsMarginV = 0
            .Image = ""
            .JustifyHotSpots = False
            .OffsetH = 0
            .OffsetV = 0
            .Spanning = tscAuto
            .Separation = 0
            
            Erase .Groups
            
            .Compile = True
            
            If mnuPreviewTBSHorizontal.Checked Then
                .Style = tscHorizonal
                .Alignment = tbacTopCenter
            Else
                .Style = tscVertical
                .Alignment = tbacTopLeft
            End If
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
        
        .UnfoldingSound.OnMouseOver = ""
        
        .FontSubstitutions = ""
        
        .DoFormsTweak = True
        .DWSupport = False
        .NS4ClipBug = False
        .OPHelperFunctions = False
        .ImageReadySupport = False
        
        .HideDelay = 200
        .SubMenusDelay = 100
        .RootMenusDelay = 100
        .SelChangeDelay = 100
        .AnimSpeed = 35
                
        .ExportHTMLParams = GenExpHTMLPref("", Project.Name, Project.FileName)
    End With
    
    Dim n As Node
    For Each n In tvRORMenus.Nodes
        If n.Bold And n.Text = cmbRoot.Text Then
            GenToolbarItems n.Child, True
            GenToolbarItems n.Child, False
            Exit For
        End If
    Next n
    
    Dim i As Integer
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            .nTextColor = vbBlack
            .nBackColor = RGB(&HD7, &HEE, &HF3)
            .CmdsFXnColor = vbWhite
            .CmdsFXNormal = cfxcRaised
            
            .hTextColor = vbBlack
            .hBackColor = RGB(230, 230, 230)
            .CmdsFXhColor = vbWhite
            .CmdsFXOver = cfxcRaised
            
            .CmdsFXSize = 1
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
            If mnuPreviewTBSHorizontal.Checked Then
                .CaptionAlignment = tacCenter
            Else
                .CaptionAlignment = tacLeft
            End If
            .IncludeInToolbar = True
            .Compile = True
            
            .CmdsMarginX = 8
            .CmdsMarginY = 3
            
            .BorderStyle = cfxcNone
            .FrameBorder = 0
            
            .iCursor.cType = iccHand
        End With
    Next i
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            .CmdsFXhColor = MenuGrps(1).CmdsFXhColor
            .CmdsFXnColor = MenuGrps(1).CmdsFXnColor
            .CmdsFXNormal = MenuGrps(1).CmdsFXNormal
            .CmdsFXOver = MenuGrps(1).CmdsFXOver
            .CmdsFXSize = MenuGrps(1).CmdsFXSize
            .CmdsMarginX = MenuGrps(1).CmdsMarginX
            .CmdsMarginY = MenuGrps(1).CmdsMarginY
            .hBackColor = MenuGrps(1).hBackColor
            .HoverFont = MenuGrps(1).DefHoverFont
            .hTextColor = MenuGrps(1).hTextColor
            .iCursor = MenuGrps(1).iCursor
            .nBackColor = MenuGrps(1).nBackColor
            .NormalFont = MenuGrps(1).DefNormalFont
            .nTextColor = MenuGrps(1).nTextColor
            
            .LeftImage = MenuGrps(1).tbiLeftImage
            .BackImage = MenuGrps(1).tbiBackImage
            '.RightImage = MenuGrps(1).tbiRightImage
        End With
    Next i
    
    GenPreview

End Sub

Private Sub GenPreview()

    Dim AddInP() As AddInParameter
    Dim ff As Integer
    
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
    
    wbPreview.Navigate AppPath + "Preview\index.html"

End Sub

Private Sub GenToolbarItems(n As Node, JustCount As Boolean)
    
    ReDim Project.Toolbars(1).Groups(0)

    Do
        Do While Not (n Is Nothing)
            If n Is Nothing Then
                Exit Do
            Else
                If n.children > 0 Then Exit Do
                Set n = n.Next
            End If
        Loop
        If n Is Nothing Then Exit Do
        
        If JustCount Then
            tGrps = tGrps + 1
        Else
            AddGrp n.Text, n.Tag, False
            If mnuPreviewTBSHorizontal.Checked Then
                MenuGrps(UBound(MenuGrps)).Alignment = gacBottomLeft
            Else
                With MenuGrps(UBound(MenuGrps))
                    .Alignment = gacRightTop
                    .tbiRightImage.NormalImage = AppPath + "exhtml\arrow_black.gif"
                    .tbiRightImage.HoverImage = AppPath + "exhtml\arrow_black.gif"
                    .tbiRightImage.w = 10
                    .tbiRightImage.h = 10
                End With
            End If
            With Project.Toolbars(1)
                ReDim Preserve .Groups(UBound(.Groups) + 1)
                .Groups(UBound(.Groups)) = MenuGrps(UBound(MenuGrps)).Name
            End With
        End If
        
        If n.children Then CreateSubMenus n.Child, UBound(MenuGrps), JustCount
        
        With MenuGrps(UBound(MenuGrps))
            If .Actions.OnMouseOver.Type <> atcCascade Then
                .tbiRightImage.NormalImage = ""
                .tbiRightImage.HoverImage = ""
            End If
        End With

        Set n = n.Next
    Loop
End Sub

Private Sub CreateSubMenus(n As Node, g As Integer, JustCount As Boolean)
    
    Dim cmd As MenuCmd
    Dim sText As String

    Do
        Do While Not (n Is Nothing)
            If n Is Nothing Then
                Exit Do
            Else
                If n.children > 0 Then Exit Do
                Set n = n.Next
            End If
        Loop
        If n Is Nothing Then Exit Do
        
        DoEvents
        
        If Not JustCount Then
            With cmd
                .Name = "C" & UBound(MenuCmds)
                .Caption = AutoWrapCaption(n.Text)
                .Parent = g
                .Compile = True
                
                SetURL n.Tag, .Actions
                sText = getDesc(Str2XML(n.Tag).firstChild)
                If sText <> .Caption Then
                    .WinStatus = Replace(sText, """", "&#34;")
                Else
                    .WinStatus = "%c"
                End If
            End With
            
            AddMenuCommand GetCmdParams(cmd), True, True, True
        End If
        
        If HasSubMenus(n) Then
            If JustCount Then
                tGrps = tGrps + 1
            Else
                AddGrp n.Text, "", True
                
                With MenuCmds(UBound(MenuCmds))
                    .RightImage.NormalImage = AppPath + "distroImages\arrow_black.gif"
                    .RightImage.HoverImage = AppPath + "distroImages\arrow_black.gif"
                    .RightImage.w = 10
                    .RightImage.h = 10
                    With .Actions.OnMouseOver
                        .Type = atcCascade
                        .TargetMenu = UBound(MenuGrps)
                        .TargetMenuAlignment = gacRightTop
                    End With
                End With
            End If
            
            CreateSubMenus n.Child, UBound(MenuGrps), JustCount
        End If
        
        Set n = n.Next
    Loop

End Sub

Private Function Str2XML(s As String) As IXMLDOMNode

    Dim d As MSXML2.DOMDocument
    
    Set d = New MSXML2.DOMDocument
    
    d.loadXML s
    
    Set Str2XML = d

End Function

Private Function HasSubMenus(ByVal n As Node) As Boolean

    Set n = n.Child
    Do While Not n Is Nothing
        If n.children Then
            HasSubMenus = True
            Exit Do
        End If
        Set n = n.Next
    Loop

End Function

Private Sub SetURL(sXML As String, ae As ActionEvents)

    If Not IsDEMO Then
        Dim sUrl As String
        sUrl = getURL(Str2XML(sXML).firstChild)
        If sUrl <> "" Then
            With ae.OnClick
                .Type = atcURL
                .url = sUrl
            End With
        End If
    End If

End Sub

Private Function AutoWrapCaption(c As String) As String

    If Len(c) > 30 Then
        Dim i As Integer
        Dim w() As String
        Dim n As String
        
        w = Split(c, " ")
        For i = 0 To UBound(w)
            If Len(n) >= Len(JoinFrom(w, i + 1)) Then
                n = Trim(n) + "<br>" + JoinFrom(w, i)
                Exit For
            Else
                n = n + w(i) + " "
            End If
        Next i
    Else
        n = c
    End If
    
    AutoWrapCaption = n

End Function

Private Function JoinFrom(a() As String, f As Integer) As String

    Dim n As String
    
    For f = f To UBound(a)
        n = n + a(f) + " "
    Next f
        
    JoinFrom = Trim(n)

End Function

Private Sub AddGrp(Title As String, sXML As String, IsASubMenu As Boolean)

    cGrp = cGrp + 1
    FloodPanel.Caption = "Creating Menu: " + Title
    FloodPanel.Value = cGrp / tGrps * 100
    DoEvents

    Dim grp As MenuGrp
    Dim sText As String
    
    With grp
        .Name = "G" & UBound(MenuGrps)
        If IsASubMenu Then
            .Caption = Title
        Else
            .Caption = AutoWrapCaption(Title)
        End If
        
        If Not IsASubMenu Then
            If sXML <> "" Then SetURL sXML, .Actions
            sText = getDesc(Str2XML(sXML).firstChild)
            If sText <> .Caption Then
                .WinStatus = Replace(sText, """", "&#34;")
            Else
                .WinStatus = "%c"
            End If
        End If
    End With
    AddMenuGroup GetGrpParams(grp), True

End Sub

Private Sub parseNode(ByVal n As IXMLDOMNode, ByVal parentNode As Node)

    Dim sn As IXMLDOMNode
    Dim newNode As Node
    Dim pNode As Node
    Dim rscOf As String
    
    If IsValidNode(n) Then
        
        tNodes = tNodes + n.childNodes.Length
        FloodPanel.Value = cNode / tNodes * 100
        
        For Each sn In n.childNodes
            cNode = cNode + 1
            If IsValidNode(sn) Then
                Select Case LCase(sn.nodeName)
                    Case "resource", "item"
                        rscOf = getResourceOf(sn)
                        If rscOf = "" Then
                            Set newNode = tvRORMenus.Nodes.Add(parentNode, tvwChild)
                        Else
                            Set pNode = GetNodeByKey("B:" + rscOf)
                            Set newNode = tvRORMenus.Nodes.Add(pNode, tvwChild)
                        End If
                        If sn.Attributes.Length > 0 Then
                            If LCase(Left(sn.Attributes(0).xml, 10)) = "rdf:about=" Then
                                newNode.key = "B:" + sn.Attributes(0).Text
                            Else
                                If LCase(Left(sn.Attributes(0).xml, 13)) = "rdf:resource=" Then
                                    newNode.key = "B:" + sn.Attributes(0).Text
                                End If
                            End If
                        End If
                        newNode.Text = getTitle(sn, True)
                        newNode.Bold = True
                        newNode.Expanded = True
                        newNode.Tag = sn.xml
                        
                        FloodPanel.Caption = "Parsing: " + newNode.Text
                    'Case "resourceof", "title", "updated", "updatePeriod","price","currency"
                    'Case Else
                    Case "url", "desc", "link", "title", "seealso"
                        Set newNode = tvRORMenus.Nodes.Add(parentNode, tvwChild, , sn.nodeName + ": " & sn.Text)
                    Case "ror:image"
                    Case "ror:seealso"
                        Set newNode = tvRORMenus.Nodes.Add(parentNode, tvwChild, , sn.nodeName + ": " & sn.Text)
                    Case Else
                        'Stop
                End Select
                If sn.hasChildNodes Then parseNode sn, newNode
            End If
        Next sn
    End If

End Sub

Private Function GetNodeByKey(k As String) As Node

    Dim n As Node

    For Each n In tvRORMenus.Nodes
        If n.key = k Then
            Set GetNodeByKey = n
            Exit For
        End If
    Next n

End Function

Private Function getType(n As IXMLDOMNode) As String

    getType = getResourceNodeValue(n, "type")

End Function

Private Function getURL(n As IXMLDOMNode) As String

    getURL = getResourceNodeValue(n, "url")
    If getURL = "" Then getURL = getResourceNodeValue(n, "link")
    If getURL = "" Then getURL = getResourceNodeValue(n, "ror:seealso")

End Function

Private Function getDesc(n As IXMLDOMNode) As String

    getDesc = getResourceNodeValue(n, "desc")

End Function

Private Function getTitle(n As IXMLDOMNode, AllowUseURL As Boolean) As String

    getTitle = getResourceNodeValue(n, "title")
    
    If (getTitle = "" Or getTitle = """""" Or getTitle = "''") And AllowUseURL Then
        getTitle = getURL(n)
        If (Left(getTitle, 1) = """" And Right(getTitle, 1) = """") Or (Left(getTitle, 1) = "'" And Right(getTitle, 1) = "'") Then
            getTitle = Mid(getTitle, 2, Len(getTitle) - 2)
        End If
        Dim p As Integer
        p = InStrRev(getTitle, "/")
        If p > 0 Then getTitle = Mid(getTitle, p + 1)
    End If
    
    getTitle = removeInvalidChars(getTitle)

End Function

Private Function getResourceNodeValue(n As IXMLDOMNode, Name As String) As String

    Dim sn As IXMLDOMNode
    
    If n Is Nothing Then Exit Function

    For Each sn In n.childNodes
        If LCase(sn.nodeName) = Name Then
            getResourceNodeValue = removeInvalidChars(sn.Text)
            Exit For
        End If
    Next sn

End Function

Private Function removeInvalidChars(ByVal s As String) As String

    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    
    removeInvalidChars = s

End Function

Private Function getResourceOf(n As IXMLDOMNode) As String

    Dim sn As IXMLDOMNode

    For Each sn In n.childNodes
        If LCase(sn.nodeName) = "resourceof" Then
            If sn.Attributes.Length > 0 Then
                If LCase(Left(sn.Attributes(0).xml, 13)) = "rdf:resource=" Then
                    getResourceOf = removeInvalidChars(sn.Attributes(0).Text)
                End If
            End If
            Exit For
        End If
    Next sn

End Function

Private Function IsValidNode(n As IXMLDOMNode) As Boolean

    IsValidNode = (n.nodeType = NODE_DOCUMENT Or n.nodeType = NODE_ELEMENT)

End Function

Private Sub cmdClose_Click()

    CloseMe

End Sub

Private Sub CloseMe()

    SavePrgSettings
    Unload Me

End Sub

Private Sub cmdSave_Click()

    SaveAsDMB

End Sub

Private Sub SaveAsDMB()

    If tvRORMenus.Nodes.Count = 0 Then
        MsgBox "There are no menus to save. Make sure you have first selected a ROR file and that the import process has generated some menus.", vbInformation + vbOKOnly, "Failed to save project"
    End If

    On Error Resume Next
    With cDlg
        .CancelError = True
        .DialogTitle = "Save DHTML Menu Builder Project"
        If InStr(FileName, "://") = 0 Then
            .InitDir = GetFilePath(FileName)
            .FileName = GetFilePath(FileName) + GetFileName(FileName, True) + ".dmb"
        Else
            .FileName = "Untitled.dmb"
        End If
        .Filter = "DHTML Menu Builder Projects|*.dmb"
        .Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
        .ShowSave
        If Err.number > 0 Or .FileName = "" Then Exit Sub
        Project.FileName = .FileName
        Project.Name = GetFileName(.FileName)
        Project.Name = Left(Project.Name, Len(Project.Name) - 4)
        Project.HasChanged = False
    End With
    On Error GoTo SaveAsDMB_Error
    
    FloodPanel.Caption = "Saving"
    
    SaveProject False
    SaveImages ""
      
    FloodPanel.Value = 0
    
    If MsgBox("Would you like to start DHTML Menu Builder and load this project?", vbQuestion + vbYesNo, "") = vbYes Then
        Shell AppPath + "dmb.exe " + Project.FileName
    End If

    On Error GoTo 0
    Exit Sub

SaveAsDMB_Error:

    MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Form.frmMain.SaveAsDMB"

End Sub

Private Sub Form_Load()

    On Error Resume Next

    AppPath = AddTrailingSlash(App.Path, "\")
    Preferences.language = "eng"

    LoadLocalizedStrings
    LoadPrgSettings
    
    wbPreview.Navigate "about:blank"
    Caption = "DHTML Menu Builder ROR Importer"
    
    picDiv.Top = tvRORMenus.Top
    picDiv.BorderStyle = 0
    
    cmbRoot.Enabled = False
    
    tmrInit.Enabled = True

End Sub

Private Sub LoadPrgSettings()

    If GetSetting("DMB", "ROR2DMB", "X", "") = "" Then
        CenterForm Me
        
        mnuPreviewTBSVertical_Click
    Else
        Left = GetSetting("DMB", "ROR2DMB", "X")
        Top = GetSetting("DMB", "ROR2DMB", "Y")
        Width = GetSetting("DMB", "ROR2DMB", "W")
        Height = GetSetting("DMB", "ROR2DMB", "H")
        picDiv.Left = GetSetting("DMB", "ROR2DMB", "D")
        Form_Resize
        
        If GetSetting("DMB", "ROR2DMB", "TBStyle") = 0 Then
            mnuPreviewTBSHorizontal_Click
        Else
            mnuPreviewTBSVertical_Click
        End If
    End If

End Sub

Private Sub SavePrgSettings()

    SaveSetting "DMB", "ROR2DMB", "X", Left
    SaveSetting "DMB", "ROR2DMB", "Y", Top
    SaveSetting "DMB", "ROR2DMB", "W", Width
    SaveSetting "DMB", "ROR2DMB", "H", Height
    SaveSetting "DMB", "ROR2DMB", "D", picDiv.Left
    
    If mnuPreviewTBSHorizontal.Checked Then
        SaveSetting "DMB", "ROR2DMB", "TBStyle", 0
    Else
        SaveSetting "DMB", "ROR2DMB", "TBStyle", 1
    End If

End Sub

Private Sub Form_Resize()

    On Error GoTo ExitSub

    With wbPreview
        .Left = (picDiv.Left + picDiv.Width)
        .Width = Width - .Left - tvRORMenus.Left * 2
        .Height = Height - .Top - cmdClose.Height - 63 * Screen.TwipsPerPixelY
        
        lblPreview.Left = .Left
        
        tvRORMenus.Height = .Height + .Top - tvRORMenus.Top - cmbRoot.Height - 4 * Screen.TwipsPerPixelY
        tvRORMenus.Width = picDiv.Left - tvRORMenus.Left * 2
        
        lblRootTB.Left = tvRORMenus.Left
        cmbRoot.Top = tvRORMenus.Top + tvRORMenus.Height + 4 * Screen.TwipsPerPixelY
        cmbRoot.Left = lblRootTB.Left + lblRootTB.Width + 3 * Screen.TwipsPerPixelX
        cmbRoot.Width = tvRORMenus.Left + tvRORMenus.Width - cmbRoot.Left
        lblRootTB.Top = cmbRoot.Top + cmbRoot.Height / 2 - lblRootTB.Height / 2
        
        picDiv.Height = tvRORMenus.Height
        
        cmdClose.Top = .Top + .Height + 10 * Screen.TwipsPerPixelY
        cmdClose.Left = .Left + .Width - cmdClose.Width
        
        cmdSave.Top = cmdClose.Top
        cmdSave.Left = cmdClose.Left - cmdSave.Width - 8 * Screen.TwipsPerPixelY
        
        picFlood.Left = tvRORMenus.Left
        picFlood.Top = cmdClose.Top + cmdClose.Height / 2 - picFlood.Height / 2
        picFlood.Width = cmdSave.Left - 8 * Screen.TwipsPerPixelX
        
        lblOpenInfo.Width = txtFileName.Width
    End With
    
ExitSub:

End Sub

Private Sub mnuFileExit_Click()

    CloseMe

End Sub

Private Sub mnuFileOpen_Click()

    OpenROR

End Sub

Private Sub mnuFileSave_Click()

    SaveAsDMB

End Sub

Private Sub mnuHelpAbout_Click()

    frmAbout.Show vbModal
    
End Sub

Private Sub mnuPreviewTBSHorizontal_Click()

    mnuPreviewTBSHorizontal.Checked = True
    mnuPreviewTBSVertical.Checked = False
    RefreshPreview

End Sub

Private Sub RefreshPreview()

    ParseROR txtFileName.Text

End Sub

Private Sub mnuPreviewTBSVertical_Click()

    mnuPreviewTBSHorizontal.Checked = False
    mnuPreviewTBSVertical.Checked = True
    RefreshPreview

End Sub

Private Sub picDiv_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsDragging = True
    mX = x

End Sub

Private Sub picDiv_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If IsDragging Then
        picDiv.Left = picDiv.Left + (x - mX)
        If picDiv.Left < cmdBrowse.Left + cmdBrowse.Width Then picDiv.Left = cmdBrowse.Left + cmdBrowse.Width
        Form_Resize
    End If

End Sub

Private Sub picDiv_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsDragging = False

End Sub

Private Sub txtFileName_KeyPress(KeyAscii As Integer)

    If KeyAscii = vbKeyReturn Then ParseROR txtFileName.Text

End Sub

Private Sub tmrInit_Timer()

          Dim i As Integer
          
10        On Error GoTo tmrInit_Timer_Error

20        tmrInit.Enabled = False
          
30        GetPrgPrefs
          
40        cSep = Chr(255) + Chr(255)
          
50        AppPath = GetSetting("DMB", "RegInfo", "InstallPath")
60        SetupTempFolders
          
70        SetTemplateDefaults
          
80        Set FloodPanel.PictureControl = picFlood
          Dim UIObjects(1 To 3) As Object
90        Set UIObjects(1) = frmMain
100       Set UIObjects(2) = FloodPanel
110       Set UIObjects(3) = frmMain
120       SetUI UIObjects
          
          Dim VarObjects(1 To 7) As Variant
130       VarObjects(1) = AppPath
140       VarObjects(2) = ""
150       VarObjects(3) = ""
160       VarObjects(4) = GetTEMPPath
170       VarObjects(5) = cSep
180       VarObjects(6) = nwdPar
190       VarObjects(7) = StatesPath
200       SetVars VarObjects
              
210       If Not FileExists(AppPath + "dmb.exe") Then
220           MsgBox "DHTML Menu Builder must be installed in order to run this Converter", vbCritical + vbOKOnly, "DHTML Menu Builder not detected"
230           CloseMe
240       End If
          
          'ParseROR "C:\Documents and Settings\Administrator\Desktop\acme.xml"
          
250       If IsDEMO Then
260           MsgBox "Although the ROR Importer can be used with the DEMO version of DHTML Menu Builder, the links in the menus will be disabled." + vbCrLf + "Please consult DHTML Menu Builder's documentation for further information regarding the DEMO version limitations.", vbCritical + vbOKOnly, "DHTML Menu Builder not registered"
270       End If

280       On Error GoTo 0
290       Exit Sub

tmrInit_Timer_Error:

300       MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Form.frmMain.tmrInit_Timer"

End Sub

Private Sub GetPrgPrefs()

    USER = GetSetting("DMB", "RegInfo", "User", "DEMO")
    COMPANY = GetSetting("DMB", "RegInfo", "Company", "DEMO")
    USERSN = GetSetting("DMB", "RegInfo", "SerialNumber", "")
    
    'If USER = "DEMO" Or USER = "" Or USERSN = "" Then End
    
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

Private Sub SetupTempFolders()

    TempPath = GetTEMPPath
    PreviewPath = AppPath + "Preview\"
    If App.PrevInstance Then
        StatesPath = AppPath + "States\" & Timer * 100 & "\"
    Else
        StatesPath = AppPath + "States\Default\"
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
