VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRsc 
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   4500
   ClientTop       =   7020
   ClientWidth     =   5130
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
   ScaleHeight     =   1680
   ScaleWidth      =   5130
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   555
      ScaleHeight     =   390
      ScaleWidth      =   330
      TabIndex        =   0
      Top             =   615
      Width           =   330
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   195
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   126
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":0000
            Key             =   "mnuMenuColor|mnuContextColor"
            Object.Tag             =   "Color"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":015C
            Key             =   "mnuMenuFont|mnuContextFont"
            Object.Tag             =   "Font"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":02B8
            Key             =   "mnuMenuCursor|mnuContextCursor"
            Object.Tag             =   "Cursor"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":0414
            Key             =   "mnuMenuImage|mnuContextImage"
            Object.Tag             =   "Image"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":0570
            Key             =   "mnuMenuSFX|mnuContextSFX"
            Object.Tag             =   "Special Effects"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":06CA
            Key             =   "btnUp"
            Object.Tag             =   "Up"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":0826
            Key             =   "btnDown"
            Object.Tag             =   "Down"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":0982
            Key             =   "mnuFileNew"
            Object.Tag             =   "New"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":0F1C
            Key             =   "mnuFileOpen"
            Object.Tag             =   "Open"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1076
            Key             =   "mnuFileSave"
            Object.Tag             =   "Save"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1610
            Key             =   "mnuEditCopy|mnuContextCopy"
            Object.Tag             =   "Copy"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1BAA
            Key             =   "mnuEditPaste|mnuContextPaste"
            Object.Tag             =   "Paste"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":2144
            Key             =   "mnuEditUndo"
            Object.Tag             =   "Undo"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":26DE
            Key             =   "mnuEditRedo"
            Object.Tag             =   "Redo"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":2C78
            Key             =   "mnuEditDelete|mnuContextDelete|mnuTBContextDelete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":3212
            Key             =   "mnuMenuAddGroup|mnuTBContextAddGroup"
            Object.Tag             =   "New Group"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":37AE
            Key             =   "mnuMenuAddCommand|mnuContextAddCommand"
            Object.Tag             =   "New Command"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":3D4A
            Key             =   "mnuMenuAddSeparator|mnuContextAddSeparator"
            Object.Tag             =   "New Separator"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":42E6
            Key             =   ""
            Object.Tag             =   "Group"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":4882
            Key             =   ""
            Object.Tag             =   "Command"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":4E1E
            Key             =   ""
            Object.Tag             =   "Separator"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":53BA
            Key             =   "mnuToolsPreview"
            Object.Tag             =   "Preview"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":5956
            Key             =   "mnuToolsHotSpotsEditor"
            Object.Tag             =   "HotSpots Editor"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":5EF2
            Key             =   "mnuFileProjProp"
            Object.Tag             =   "Properties"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":628C
            Key             =   "mnuToolsCompile"
            Object.Tag             =   "Compile"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":70DE
            Key             =   "mnuHelpUpgrade"
            Object.Tag             =   "Upgrade"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":7686
            Key             =   "mnuHelpXFXFAQ"
            Object.Tag             =   "FAQ"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":79DA
            Key             =   "mnuHelpXFXNews"
            Object.Tag             =   "News"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":7D2E
            Key             =   "mnuHelpXFXSupport"
            Object.Tag             =   "Support"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":8082
            Key             =   ""
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":83D6
            Key             =   "mnuHelpXFXPublicForum"
            Object.Tag             =   "Forum"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":872A
            Key             =   "mnuEditFind"
            Object.Tag             =   "Find"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":8884
            Key             =   "NoEvents"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":8E1E
            Key             =   "OverCascade"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":93B8
            Key             =   "Click"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":9952
            Key             =   "ClickCascade"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":9EEC
            Key             =   "DoubleClick"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":A486
            Key             =   "DoubleClickCascade"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":AA20
            Key             =   "Over"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":AFBA
            Key             =   "Disabled"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":B554
            Key             =   "mnuMenuMargins|mnuContextMargins"
            Object.Tag             =   "Margins"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":B8EE
            Key             =   "mnuToolsPublish"
            Object.Tag             =   "Publish"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":E0A0
            Key             =   "mnuMenuSound"
            Object.Tag             =   "Sounds"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":E1FA
            Key             =   "GClick"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":E794
            Key             =   "GOver"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":ED2E
            Key             =   "GOverCascade"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":F2C8
            Key             =   "GNoEvents"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":F862
            Key             =   "GClickCascade"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":FDFC
            Key             =   "GDoubleClick"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":10396
            Key             =   "GDoubleClickCascade"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":10930
            Key             =   "GDisabled"
            Object.Tag             =   "DisabledGroup"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":10ECA
            Key             =   "mnuEditRename|mnuContextRename|mnuTBContextRename"
            Object.Tag             =   "Rename"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":11264
            Key             =   "mnuEditPreferences"
            Object.Tag             =   "Preferences"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":115FE
            Key             =   "mnuFileRF"
            Object.Tag             =   "DMBProjectIcon"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":14308
            Key             =   "EmptyIcon"
            Object.Tag             =   "EmptyIcon"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":146A2
            Key             =   "mnuToolsAddInEditor"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":14A3C
            Key             =   ""
            Object.Tag             =   "Left"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":14B96
            Key             =   ""
            Object.Tag             =   "Right"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":14CF0
            Key             =   ""
            Object.Tag             =   "Normal"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":14E4A
            Key             =   ""
            Object.Tag             =   "Over"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":14FA4
            Key             =   ""
            Object.Tag             =   "Font Bold"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":150B6
            Key             =   ""
            Object.Tag             =   "Font Italic"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":151C8
            Key             =   ""
            Object.Tag             =   "Font Underline"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":152DA
            Key             =   ""
            Object.Tag             =   "Size"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":153EC
            Key             =   ""
            Object.Tag             =   "Font Name"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":154FE
            Key             =   ""
            Object.Tag             =   "Toolbar Item"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":15658
            Key             =   ""
            Object.Tag             =   "Target Frame"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":157B2
            Key             =   ""
            Object.Tag             =   "Leading"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1590C
            Key             =   ""
            Object.Tag             =   "Group Alignment"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":15EA6
            Key             =   ""
            Object.Tag             =   "Caption Alignment"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":15FB8
            Key             =   ""
            Object.Tag             =   "Events"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":16112
            Key             =   ""
            Object.Tag             =   "Frame"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1626C
            Key             =   ""
            Object.Tag             =   "Line"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":16806
            Key             =   ""
            Object.Tag             =   "Highlight Effects"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":16DA0
            Key             =   ""
            Object.Tag             =   "Transparency"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":16EFA
            Key             =   ""
            Object.Tag             =   "Shadow"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":17054
            Key             =   ""
            Object.Tag             =   "All Properties"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":171AE
            Key             =   ""
            Object.Tag             =   "EventOver"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":17308
            Key             =   ""
            Object.Tag             =   "EventDoubleClick"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":17462
            Key             =   ""
            Object.Tag             =   "EventClick"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":175BC
            Key             =   ""
            Object.Tag             =   "URL"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":17B56
            Key             =   ""
            Object.Tag             =   "Action Type"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":180F0
            Key             =   ""
            Object.Tag             =   "Border Size"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1848A
            Key             =   ""
            Object.Tag             =   "Command Horizontal Margin"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":18824
            Key             =   ""
            Object.Tag             =   "Command Vertical Margin"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":18BBE
            Key             =   ""
            Object.Tag             =   "Vertical Margin"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":18F58
            Key             =   ""
            Object.Tag             =   "Horizontal Margin"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":192F2
            Key             =   ""
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1968C
            Key             =   ""
            Object.Tag             =   "New Window"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":19A26
            Key             =   ""
            Object.Tag             =   "Colored Borders"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":19DC0
            Key             =   ""
            Object.Tag             =   "Group Width"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1A15A
            Key             =   ""
            Object.Tag             =   "Group Height"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1A4F4
            Key             =   ""
            Object.Tag             =   "Text Color"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1AA8E
            Key             =   ""
            Object.Tag             =   "Back Color"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1B028
            Key             =   ""
            Object.Tag             =   "Status Text"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1B5C2
            Key             =   ""
            Object.Tag             =   "Caption"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1BB5C
            Key             =   "mnuRegisterUnlock"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1C836
            Key             =   ""
            Object.Tag             =   "Commands Layout"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1CBD0
            Key             =   ""
            Object.Tag             =   "Group Effects"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1D16A
            Key             =   ""
            Object.Tag             =   "Overlay"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1D704
            Key             =   ""
            Object.Tag             =   "HS-Text"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1DC9E
            Key             =   ""
            Object.Tag             =   "HS-Image"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1E238
            Key             =   ""
            Object.Tag             =   "HS-DynaText"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1E7D2
            Key             =   "mnuHelpContents"
            Object.Tag             =   "Help"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1ED6C
            Key             =   "www"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1F306
            Key             =   "mnuMenuToolbarProperties|mnuToolsToolbarsEditor|mnuTBContextToolbarProperties"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1F460
            Key             =   "mnuToolsLCMan|mnuToolsLC"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1F5BA
            Key             =   "mnuMenuAddSubGroup|mnuContextAddSubGroup"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":1FB54
            Key             =   "mnuMenuLength|mnuContextLength"
            Object.Tag             =   "SeparatorLength"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":200EE
            Key             =   ""
            Object.Tag             =   "Justify"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":20688
            Key             =   ""
            Object.Tag             =   "Toolbar Alignment"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":20C22
            Key             =   ""
            Object.Tag             =   "Spanning"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":20D7C
            Key             =   ""
            Object.Tag             =   "Follow Scrolling"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":20ED6
            Key             =   ""
            Object.Tag             =   "Toolbar Offset"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":21030
            Key             =   "mnuMenuAddToolbar|mnuTBContextAddToolbar"
            Object.Tag             =   "Add Toolbar"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":2118A
            Key             =   "mnuMenuRemoveToolbar"
            Object.Tag             =   "Remove Toolbar"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":212E4
            Key             =   "mnuEditFindReplace"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":2187E
            Key             =   "mnuHelpSearch"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":21E18
            Key             =   "mnuHelpXFXHomePage"
            Object.Tag             =   "xFXLogo"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":221B2
            Key             =   ""
            Object.Tag             =   "Scrolling"
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":2254C
            Key             =   "mnuFileNewEmpty"
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":22AE6
            Key             =   "mnuFileNewFromDir"
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":23080
            Key             =   "mnuFileNewFromPreset"
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":2361A
            Key             =   "mnuFileNewFromWizard"
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":23BB4
            Key             =   "Transparent"
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRsc.frx":23D0E
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRsc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

