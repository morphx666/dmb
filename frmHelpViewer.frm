VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmHelpViewer 
   Caption         =   "DHTML Menu Builder Help"
   ClientHeight    =   6000
   ClientLeft      =   2025
   ClientTop       =   4470
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelpViewer.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   9585
   Begin SHDocVwCtl.WebBrowser wbCtrl 
      Height          =   4395
      Left            =   6195
      TabIndex        =   7
      Top             =   330
      Width           =   4590
      ExtentX         =   8096
      ExtentY         =   7752
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
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
   Begin VB.PictureBox picDiv 
      BorderStyle     =   0  'None
      Height          =   5205
      Left            =   3555
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5205
      ScaleWidth      =   570
      TabIndex        =   9
      Top             =   390
      Width           =   570
   End
   Begin MSComctlLib.ImageList ilONIcons 
      Left            =   870
      Top             =   5025
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":27A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":289C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":2996
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":2A90
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":2D26
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":2FBC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":3252
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":34E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":377E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":3A14
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":3CAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":3F40
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbBrowser 
      Height          =   285
      Left            =   2580
      TabIndex        =   14
      Top             =   30
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   503
      ButtonWidth     =   529
      ButtonHeight    =   503
      ToolTips        =   0   'False
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Style           =   1
      ImageList       =   "ilIcons"
      HotImageList    =   "ilONIcons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbShowHide"
            Object.ToolTipText     =   "Show/Hide Contents"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbBack"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbFwd"
            Object.ToolTipText     =   "Forward"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbHome"
            Object.ToolTipText     =   "Introduction"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbFind"
            Object.ToolTipText     =   "Find"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbPrint"
            Object.ToolTipText     =   "Print"
            ImageIndex      =   8
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   165
      Top             =   5055
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   13
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":449E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":4598
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":4692
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":478C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":4A22
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":4CB8
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":4F4E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":51E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":547A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":54FB
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":5A59
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmHelpViewer.frx":5CEF
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picSearch 
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   600
      ScaleHeight     =   4695
      ScaleWidth      =   4650
      TabIndex        =   11
      Top             =   720
      Visible         =   0   'False
      Width           =   4650
      Begin VB.CommandButton cmdHighlight 
         Caption         =   "Highlight"
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
         Height          =   345
         Left            =   3600
         TabIndex        =   15
         Top             =   615
         Width           =   870
      End
      Begin VB.TextBox txtKeywords 
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
         Left            =   15
         TabIndex        =   1
         Top             =   210
         Width           =   3465
      End
      Begin VB.CommandButton cmdSearch 
         Caption         =   "&Search"
         Default         =   -1  'True
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
         Height          =   345
         Left            =   3600
         TabIndex        =   2
         Top             =   180
         Width           =   870
      End
      Begin VB.OptionButton opSOptions 
         Caption         =   "Match any word"
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
         Index           =   0
         Left            =   240
         TabIndex        =   3
         Top             =   630
         Value           =   -1  'True
         Width           =   1785
      End
      Begin VB.OptionButton opSOptions 
         Caption         =   "Match all words"
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
         Left            =   240
         TabIndex        =   4
         Top             =   855
         Width           =   1785
      End
      Begin VB.OptionButton opSOptions 
         Caption         =   "Match exact phrase"
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
         Index           =   2
         Left            =   240
         TabIndex        =   5
         Top             =   1095
         Width           =   1785
      End
      Begin MSComctlLib.ListView lvIndex 
         Height          =   2340
         Left            =   15
         TabIndex        =   6
         Top             =   1545
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   4128
         SortKey         =   1
         View            =   3
         LabelEdit       =   1
         SortOrder       =   -1  'True
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         MousePointer    =   99
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MouseIcon       =   "frmHelpViewer.frx":624D
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "pTitle"
            Text            =   "Page Title"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Key             =   "pHits"
            Text            =   "Hits"
            Object.Width           =   2540
         EndProperty
      End
      Begin xfxLine3D.ucLine3D uc3DLineCtrl 
         Height          =   30
         Left            =   15
         TabIndex        =   12
         Top             =   1395
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   53
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Keywords"
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
         Left            =   15
         TabIndex        =   13
         Top             =   -15
         Width           =   705
      End
   End
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   10
      Top             =   5715
      Width           =   9585
      _ExtentX        =   16907
      _ExtentY        =   503
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
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
   Begin MSComctlLib.TreeView tvContents 
      Height          =   4155
      Left            =   45
      TabIndex        =   0
      Top             =   405
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   7329
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ilIcons"
      Appearance      =   1
      MousePointer    =   99
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "frmHelpViewer.frx":63AF
   End
   Begin MSComctlLib.TabStrip tsContents 
      Height          =   4905
      Left            =   0
      TabIndex        =   8
      Top             =   30
      Width           =   3390
      _ExtentX        =   5980
      _ExtentY        =   8652
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Contents"
            Key             =   "tsContents"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "tsSearch"
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
Attribute VB_Name = "frmHelpViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsDrag As Boolean
Dim cx As Long
Dim SelHelpFile As String
Dim AppPath As String
Dim ff As Integer
Dim HideContents As Boolean
Dim LastDivLeft As Long
Dim ServerPath As String

Friend Sub ForceSelection(FileName As String)

    If IsExternalLink(FileName) Then
        wbCtrl.Navigate FileName
    Else
        wbCtrl.Navigate AppPath + "Help\" + FileName
    End If

End Sub

Private Sub cmdHighlight_Click()

    HighlightMatches

End Sub

Private Sub Form_Load()

    If Val(GetSetting("DMB", "HelpViewerWinPos", "X")) = 0 Then
        CenterForm Me
        picDiv.Move 3555, 330, 60, Height
    Else
        Top = GetSetting("DMB", "HelpViewerWinPos", "X")
        Left = GetSetting("DMB", "HelpViewerWinPos", "Y")
        Width = GetSetting("DMB", "HelpViewerWinPos", "W")
        Height = GetSetting("DMB", "HelpViewerWinPos", "H")
        WindowState = Val(GetSetting("DMB", "HelpViewerWinPos", "State"))
        picDiv.Move Val(GetSetting("DMB", "HelpViewerWinPos", "S")), 330, 60, Height
    End If
    
    SelHelpFile = CStr(Command)
    AppPath = GetSetting("DMB", "RegInfo", "InstallPath")

    CreateContents (GetSetting("DMB", "RegInfo", "SubSystem") = "LITE")
    
    If SelHelpFile = "" Then
        ForceSelection "introduction.htm"
    Else
        ForceSelection SelHelpFile
    End If

End Sub

Private Sub CreateContents(ByVal IsLite As Boolean)

    Dim nItem As Node

    With tvContents.Nodes
        .Add , , "introduction", "Introduction", 2, 3
            .Add "introduction", tvwChild, "getting_started", "Getting Started", 1, 1
            .Add "introduction", tvwChild, "preparing_your_web_site", "Preparing your Web Site", 1, 1
        
        .Add , , "project_properties", "Project Properties", 2, 3
            .Add "project_properties", tvwChild, "configurations", "Configurations", 1, 1
        
        .Add , , "creating_menus", "Creating Menus", 2, 3
            .Add "creating_menus", tvwChild, "groups", "Groups", 1, 1
            .Add "creating_menus", tvwChild, "commands", "Commands", 1, 1
            .Add "creating_menus", tvwChild, "creating_toolbars", "Creating Toolbars", 2, 3
                .Add "creating_toolbars", tvwChild, "tb_setstyle", "Setting the Toolbars Style", 1, 1
                .Add "creating_toolbars", tvwChild, "tb_pos", "Positioning the Toolbars", 1, 1
            .Add "creating_menus", tvwChild, "menuanatomy", "Anatomy of a Menu", 1, 1
        
        .Add , , "working_styles", "Working With Styles", 2, 3
            .Add "working_styles", tvwChild, "copypaste_styles", "Copying and Pasting Styles", 1, 1
            .Add "working_styles", tvwChild, "scopeselector", "The Scope Selector", 1, 1
            .Add "working_styles", tvwChild, "previewing", "Previewing", 1, 1
        
        .Add , , "presets", "Presets", 2, 3
            .Add "presets", tvwChild, "presets_new", "Creating Projects from Presets", 1, 1
            .Add "presets", tvwChild, "presets_applying", "Applying Styles from Presets", 1, 1
            .Add "presets", tvwChild, "presets_creating", "Creating Presets", 1, 1
            .Add "presets", tvwChild, "presets_sharing", "Sharing Presets", 1, 1
        
        .Add , , "implementing_menus", "Implementing the Menus", 2, 3
            If Not IsLite Then
                .Add "implementing_menus", tvwChild, "working_with_frames", "Working with Frames", 1, 1
                .Add "implementing_menus", tvwChild, "working_with_frames_adv", "Advanced Frames Implementation", 1, 1
                .Add "implementing_menus", tvwChild, "working_with_iframes", "Working with IFrames", 1, 1
            End If
            .Add "implementing_menus", tvwChild, "installingmenus", "Installing Menus", 1, 1
            .Add "implementing_menus", tvwChild, "hotspot_editor", "The HotSpots Editor", 1, 1
            If Not IsLite Then
                .Add "implementing_menus", tvwChild, "share_projects", "Multiple Projects", 1, 1
            End If
        
        .Add , , "tutorial", "Tutorials", 2, 3
            .Add "tutorial", tvwChild, "thetutorial_tb_01", "Using a Toolbar", 1, 1
            .Add "tutorial", tvwChild, "thetutorial_ptb_01", "Positioning the menus", 2, 3
                .Add "thetutorial_ptb_01", tvwChild, "thetutorial_ptb_02", "A horizontal menu must be placed inside a horizontal cell", 1, 1
                .Add "thetutorial_ptb_01", tvwChild, "thetutorial_ptb_04", "A vertical menu must be placed inside a left aligned vertical cell", 1, 1
                .Add "thetutorial_ptb_01", tvwChild, "thetutorial_ptb_07", "A vertical menu must be placed inside a cell where its left position can vary depending on the browser's window size", 1, 1
                .Add "thetutorial_ptb_01", tvwChild, "thetutorial_ptb_10", "When the menus appear perfectly positioned under some browsers and not others", 1, 1
                
                .Add "thetutorial_ptb_01", tvwChild, "http://software.xfx.net/utilities/dmbuilder/content/tipstricks/tbpos.php", "Understading the Positioning Controls", 10, 10
                .Add "thetutorial_ptb_01", tvwChild, "http://software.xfx.net/utilities/dmbuilder/content/tipstricks/inacell.php", "Placing a toolbar inside a table's cell", 10, 10
            .Add "tutorial", tvwChild, "thetutorial_ntb_01", "Using Hotspots", 1, 1
        
        .Add , , "more_help", "More Help Resources", 2, 3
            .Add "more_help", tvwChild, "http://software.xfx.net/utilities/dmbuilder/faq/index.html", "Frequently Asked Questions", 10, 10
            .Add "more_help", tvwChild, "http://software.xfx.net/utilities/dmbuilder/tipstricks/index.html", "Tips & Tricks", 10, 10
            .Add "more_help", tvwChild, "http://software.xfx.net/uboards/uboard_dmb.htm", "Public Forum", 10, 10
            .Add "more_help", tvwChild, "file:///" + App.Path + "\Samples", "Sample Projects", 1, 1
            .Add "more_help", tvwChild, "/dialogs/index", "Dialogs", 1, 1
        
        .Add , , "add_info", "Additional Information", 2, 3
            .Add "add_info", tvwChild, "seo", "Search Engine Optimizations", 1, 1
            .Add "add_info", tvwChild, "browsers_limitations", "Incompatibilities and limitations", 1, 1
            .Add "add_info", tvwChild, "demo", "DEMO version limitations", 1, 1
            If IsLite Then
                .Add "add_info", tvwChild, "http://software.xfx.net/utilities/dmbuilderlite/stdvsse.php", "LITE version limitations", 10, 10
            End If
            .Add "add_info", tvwChild, "importing", "Importing Projects from older versions", 1, 1
            .Add "add_info", tvwChild, "reginfo", "Registration Information", 1, 1
            .Add "add_info", tvwChild, "privacyinfo", "Privacy Information", 1, 1
            .Add "add_info", tvwChild, "whatsnew", "What's New", 9, 9
            .Add "add_info", tvwChild, "ack", "Acknowledgements", 1, 1
            .Add "add_info", tvwChild, "http://software.xfx.net/utilities/dmbuilder/webslist.php", "Customers' Web Sites", 10, 10
            .Add "add_info", tvwChild, "glossary", "Glossary", 1, 1
        
        If Not IsLite Then
            .Add , , "adv_topics", "Advanced Topics", 2, 3
                .Add "adv_topics", tvwChild, "tb_vc", "Toolbars' Visibility Condition", 1, 1
                .Add "adv_topics", tvwChild, "aie", "AddIn Editor", 1, 1
        End If
    End With
    
    For Each nItem In tvContents.Nodes
        nItem.Expanded = False
        If nItem.SelectedImage = 3 Then
            nItem.SelectedImage = 2
            nItem.ExpandedImage = 3
        End If
    Next nItem

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If WindowState = vbNormal Then
        SaveSetting "DMB", "HelpViewerWinPos", "X", Top
        SaveSetting "DMB", "HelpViewerWinPos", "Y", Left
        SaveSetting "DMB", "HelpViewerWinPos", "W", Width
        SaveSetting "DMB", "HelpViewerWinPos", "H", Height
    End If
    SaveSetting "DMB", "HelpViewerWinPos", "S", IIf(HideContents, LastDivLeft, picDiv.Left)
    SaveSetting "DMB", "HelpViewerWinPos", "State", WindowState
    
'    If StartMode = ActiveEXE Then
'        If Not ClassIsClosing Then
'            Me.Visible = False
'            Cancel = 1
'        End If
'    End If

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    If WindowState = vbMinimized Then Exit Sub
    
    If HideContents Then
        picDiv.Left = 0
        picSearch.Visible = False
    Else
        If picDiv.Left < 1635 Then picDiv.Left = 1635
        If Width - picDiv.Left < 2235 Then picDiv.Left = Width - 2235
        picSearch.Visible = (tsContents.SelectedItem.Key = "tsSearch")
    End If

    tsContents.Move 0, 30, picDiv.Left - 30, Height - 750
    tvContents.Move 45, 405, tsContents.Width - 120, tsContents.Height - 450
    
    wbCtrl.Move picDiv.Left + 90, 330, Width - picDiv.Left - 240, Height - 1050
    
    tbBrowser.Left = wbCtrl.Left
    
    picDiv.Height = tsContents.Height - 300
    
    '------------
    
    picSearch.Move 45, 405, tsContents.Width - 120, tsContents.Height - 450
    txtKeywords.Width = picSearch.Width - cmdSearch.Width - 180
    cmdSearch.Left = picSearch.Width - cmdSearch.Width - 60
    cmdHighlight.Left = cmdSearch.Left
    
    uc3DLineCtrl.Width = picSearch.Width
    
    lvIndex.Width = picSearch.Width - 15
    lvIndex.Height = picSearch.Height - lvIndex.Top - 15
    
End Sub

Private Sub lvIndex_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    Dim nItem As ListItem
    
    For Each nItem In lvIndex.ListItems
        nItem.SubItems(1) = Format(nItem.SubItems(1), "0000")
    Next nItem

    With lvIndex
        If .SortKey = ColumnHeader.Index - 1 Then
            If .SortOrder = lvwAscending Then
                .SortOrder = lvwDescending
            Else
                .SortOrder = lvwAscending
            End If
        Else
            .SortKey = ColumnHeader.Index - 1
            .SortOrder = lvwAscending
        End If
        .Sorted = True
    End With
    
    For Each nItem In lvIndex.ListItems
        nItem.SubItems(1) = CStr(Val(nItem.SubItems(1)))
    Next nItem

End Sub

Private Sub picDiv_DblClick()

    Dim ch As ColumnHeader
    Dim acc As Integer
    
    acc = lvIndex.Left + picDiv.Width + 360
    
    Select Case tsContents.SelectedItem.Key
        Case "tsContents"
            
        Case "tsSearch"
            For Each ch In lvIndex.ColumnHeaders
                acc = acc + ch.Width
            Next ch
            picDiv.Left = acc
    End Select
    
    Form_Resize
    Refresh
    
End Sub

Private Sub picDiv_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    IsDrag = True
    cx = x
    
    picDiv.BackColor = vbBlack

End Sub

Private Sub picDiv_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If IsDrag Then
        picDiv.Left = picDiv.Left + (x - cx)
        Form_Resize
        Refresh
    End If

End Sub

Private Sub picDiv_MouseUp(Button As Integer, Shift As Integer, x As Single, Y As Single)

    IsDrag = False
    
    picDiv.BackColor = Me.BackColor

End Sub

Private Sub tbBrowser_ButtonClick(ByVal Button As MSComctlLib.Button)

    On Error Resume Next

    Select Case Button.Key
        Case "tbBack"
            wbCtrl.GoBack
        Case "tbHome"
            tvContents_NodeClick tvContents.Nodes(1)
        Case "tbFwd"
            wbCtrl.GoForward
        Case "tbFind"
            SetFocus2Browser
            SendKeys "^f", True
        Case "tbPrint"
            SetFocus2Browser
            SendKeys "^p", True
        Case "tbShowHide"
            ToggleContents
    End Select

End Sub

Private Sub ToggleContents()
    
    If Not HideContents Then
        LastDivLeft = picDiv.Left
    Else
        picDiv.Left = LastDivLeft
    End If

    HideContents = Not HideContents
            
    tsContents.Visible = Not HideContents
    tvContents.Visible = Not HideContents
    
    tbBrowser.Buttons("tbShowHide").Image = IIf(HideContents, 12, 11)
    
    Form_Resize

End Sub

Private Sub SetFocus2Browser()

    On Error Resume Next
    
    DoEvents

    wbCtrl.SetFocus
    wbCtrl.Document.parentWindow.focus

End Sub

Friend Sub tsContents_Click()

    On Error Resume Next

    picSearch.Visible = (tsContents.SelectedItem.Key = "tsSearch")
    If picSearch.Visible Then txtKeywords.SetFocus
    
    If picSearch.Visible And HideContents Then
        ToggleContents
    End If

End Sub

Private Sub tvContents_NodeClick(ByVal Node As MSComctlLib.Node)

    Static LastNode As String
    Dim sDoc As String
    
    On Error GoTo ReportError
    
    If LastNode = Node.Key Then
        Node.Expanded = Not Node.Expanded
    Else
        Node.Expanded = True
        LastNode = Node.Key
    End If

    If Left(Node.Key, 7) <> "http://" And Left(Node.Key, 7) <> "file://" Then
        sDoc = Node.Key
        If InStr(sDoc, "|") Then sDoc = Split(sDoc, "|")(0)
        sDoc = AppPath + "Help\" + Replace(sDoc, "/", "\") + ".htm"
    Else
        sDoc = Node.Key
    End If
    
    sDoc = Replace(sDoc, "\\", "\")
    If FileExists(sDoc) Or IsExternalLink(sDoc) Then wbCtrl.Navigate sDoc
    
    Exit Sub
    
ReportError:
    MsgBox "Error " & Err.number & ": " & Err.Description, vbCritical + vbOKOnly

End Sub

Private Sub txtKeywords_GotFocus()

    With txtKeywords
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Private Sub txtKeywords_LinkOpen(Cancel As Integer)

    Stop

End Sub

Private Sub wbCtrl_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    If Left(URL, 7) = "http://" Or Left(URL, 6) = "ftp://" Then
        If Not IsOnline Then
            MsgBox "You have clicked on an external link." + vbCrLf + vbCrLf + _
                    "Please connect to the Internet so you can access the remote web site:" + vbCrLf + _
                    URL, vbInformation + vbOKOnly, "Unable to access external link"
            Cancel = True
        End If
    End If

End Sub

Private Sub wbCtrl_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    Dim nItem As Node
    Dim FileName As String
    Dim ItemFileName As String
    
    If InStr(URL, "\dialogs\") Then
        Set nItem = tvContents.Nodes("/dialogs/index")
    Else
        FileName = GetFileName(URL)
        For Each nItem In tvContents.Nodes
            ItemFileName = GetFileName(LCase(nItem.Key + ".htm"))
            If ItemFileName = "" Then ItemFileName = LCase(nItem.Key + ".htm")
            If ItemFileName = FileName Then
                Exit For
            End If
        Next nItem
    End If
    
    If Not nItem Is Nothing Then
        nItem.EnsureVisible
        nItem.Selected = True
        If Not nItem.Parent Is Nothing Then
            nItem.Parent.Expanded = True
        End If
    End If
    
    If Me.Visible Then
        'Me.Caption = "DHTML Menu Builder Help" + " - " + wbCtrl.Document.Title
        SetFocus2Browser
    End If
    
End Sub

Private Sub HighlightMatches()

    Dim i As Integer
    Dim k() As String
    
    If InStr(txtKeywords.Text, " ") Then
        k = Split(txtKeywords.Text, " ")
    Else
        ReDim k(0)
        k(0) = txtKeywords.Text
    End If

    If opSOptions(0).Value Then
        For i = 0 To UBound(k)
            DoHighlight k(i)
        Next i
    ElseIf opSOptions(1).Value Then
        For i = 0 To UBound(k)
            DoHighlight k(i)
        Next i
    Else
        DoHighlight txtKeywords.Text
    End If

End Sub

Private Sub DoHighlight(ByVal sStr As String)

    Dim i As Long
    Dim HTML As String
    Dim s As String
    Dim ln As Integer
    Dim j As Long
    Dim lsStr As String
    Dim m As String
    
    ln = Len(sStr)
    lsStr = LCase(sStr)
    
    HTML = wbCtrl.Document.body.innerHTML
    For i = 1 To Len(HTML) - ln
        s = Mid(HTML, i, 1)
        If s = "<" Then
            i = InStr(i, HTML, ">")
        Else
            Do
                s = Mid(HTML, i, ln)
                j = InStr(s, "<")
                If j > 0 Then
                    i = InStr(i + j, HTML, ">")
                    Exit Do
                Else
                    If LCase(s) = lsStr Then
                        m = LCase(Mid(HTML, i - 1, 1))
                        If Not (m >= "a" And m <= "z") Then
                            m = "<span style=""background-color: #0000FF; color: #FFFFFF;"">" + s + "</span>"
                            HTML = Left(HTML, i - 1) + m + Mid(HTML, i + ln)
                            i = i + Len(m)
                        End If
                    End If
                End If
                i = i + 1
            Loop
        End If
    Next i
    wbCtrl.Document.body.innerHTML = HTML

End Sub

Private Sub cmdSearch_Click()

    Dim nItm As ListItem
    Dim nRes As Integer

    lvIndex.ListItems.Clear
    lvIndex.Sorted = True
    
    If InStr(LCase(txtKeywords.Text), "easter egg") > 0 Then
        MsgBox "Nice try...", vbOKOnly, "How to get the Easter Egg"
        Exit Sub
    End If
    
    nRes = DoSearch(AppPath + "Help\")
    nRes = nRes + DoSearch(AppPath + "Help\Dialogs\")
    
    If lvIndex.ListItems.Count > 0 Then
        lvIndex.Sorted = False
        lvIndex.Sorted = True
        lvIndex.Sorted = False
        For Each nItm In lvIndex.ListItems
            nItm.SubItems(1) = CStr(Val(nItm.SubItems(1)))
        Next nItm
        
        lvIndex.ListItems(1).EnsureVisible
        lvIndex.ListItems(1).Selected = True
        lvIndex.SetFocus
    End If
    
    sbInfo.SimpleText = lvIndex.ListItems.Count & " matches total from " & nRes & " available files"
    
    CoolListView lvIndex
    
    cmdHighlight.Enabled = lvIndex.ListItems.Count > 0

End Sub

Private Function DoSearch(hlpPath As String) As Integer

    Dim k() As String
    Dim cFile As String
    Dim fCont As String
    Dim mCount As Integer
    Dim f As Integer
    Dim nItm As ListItem
    Dim fContNoHTML As String
    Dim pt As String
    
    On Error Resume Next
    
    If InStr(txtKeywords.Text, " ") = 0 Then
        ReDim k(0)
        k(0) = LCase(txtKeywords.Text)
    Else
        k = Split(LCase(txtKeywords.Text), " ")
    End If
    
    cFile = Dir(hlpPath + "*.htm*")
    While cFile <> ""
        Select Case LCase(cFile)
        Case "index.html", "contents.htm"
            'just skip these files...
        Case Else
            fCont = ReadFile(hlpPath + cFile)
            pt = PageTitle(fCont)
            
            sbInfo.SimpleText = "Searching: " + pt + " (" + cFile + ")"
            DoEvents
            
            fContNoHTML = RemoveHTMLCode(fCont)
            
            If opSOptions(0).Value Then
                mCount = MatchAny(fContNoHTML, k)
            ElseIf opSOptions(1).Value Then
                mCount = MatchAll(fContNoHTML, k)
            Else
                mCount = MatchPhrase(fContNoHTML)
            End If
            
            If mCount > 0 Then
                Err.Clear
                Set nItm = lvIndex.ListItems.Add(, cFile, pt)
                If Err.number = 0 Then
                    nItm.SubItems(1) = Format(mCount, "0000")
                End If
            End If
            
            f = f + 1
        End Select
        cFile = Dir
    Wend
    
    DoSearch = f

End Function

Private Function MatchPhrase(fCont As String) As Integer

    Dim p As Long
    Dim c As Integer

    p = 0
    Do While True
        p = InStr(p + 1, LCase(fCont), LCase(txtKeywords.Text))
        If p = 0 Then Exit Do
        c = c + 1
    Loop
        
    MatchPhrase = c

End Function

Private Function MatchAll(fCont As String, k() As String) As Integer

    Dim i As Integer
    Dim p() As Long
    Dim c As Integer
    
    ReDim p(UBound(k))
    For i = 0 To UBound(p)
        p(i) = 0
    Next i
    
    Do While True
        For i = 0 To UBound(k)
            p(i) = InStr(p(i) + 1, LCase(fCont), LCase(k(i)))
        Next i
        For i = 0 To UBound(p)
            If p(i) = 0 Then Exit Do
        Next i
        c = c + 1
    Loop
    
ReturnResults:
    MatchAll = c

End Function

Private Function MatchAny(fCont As String, k() As String) As Integer

    Dim i As Integer
    Dim p As Long
    Dim c As Integer
    
    For i = 0 To UBound(k)
        p = InStr(LCase(fCont), LCase(k(i)))
        While p > 0
            c = c + 1
            p = InStr(p + 1, LCase(fCont), LCase(k(i)))
        Wend
    Next i
    
    MatchAny = c

End Function

Private Function PageTitle(fCont As String) As String

    Dim p1 As Integer
    Dim p2 As Integer
    
    On Error Resume Next
    
    p1 = InStr(LCase(fCont), "<title>") + 7
    p2 = InStr(LCase(fCont), "</title>")
    
    PageTitle = Mid(fCont, p1, p2 - p1)

End Function

Private Function ReadFile(FileName As String) As String

    On Error Resume Next
    ff = FreeFile
    Open FileName For Input As #ff
        ReadFile = Input$(LOF(ff), ff)
    Close #ff

End Function

Private Sub lvIndex_ItemClick(ByVal item As MSComctlLib.ListItem)

    item.EnsureVisible
    item.Selected = True
    
    If Left(item.Text, 9) = "Dialogs /" Then
        wbCtrl.Navigate AppPath + "Help\Dialogs\" + item.Key
        cmdHighlight.Enabled = False
    Else
        wbCtrl.Navigate AppPath + "Help\" + item.Key
        cmdHighlight.Enabled = True
    End If
    
End Sub

Private Sub txtKeywords_Change()

    cmdSearch.Enabled = (txtKeywords.Text <> "")

End Sub

Public Function RemoveHTMLCode(ByVal sCode As String) As String

    Dim p1 As Integer
    Dim p2 As Integer
    Dim p3 As Integer
    Dim p4 As Integer
    Dim s As String
    Dim ss As String
    Dim k As Integer
    
    On Error GoTo AbortFcn
    
    s = sCode
    
    While (InStr(s, "<") > 0) And (InStr(s, ">") > 0) And k < 100
        k = k + 1
        ss = ""
        p1 = 1
        Do
            p1 = InStr(p1, s, "<")
            If p1 > 0 Then
                If p1 > 1 Then ss = ss + Left(s, p1 - 1)
                p2 = InStr(p1, s, ">")
                p3 = InStr(p1, s, " ")
                If p2 = 0 Then Exit Do
                If p3 < p2 And p3 <> 0 Then
                    p4 = p3
                    s = Replace(s, "</" + Mid(s, p1 + 1, p4 - p1 - 1) + ">", "")
                Else
                    p4 = p2
                    s = Replace(s, "</" + Mid(s, p1 + 1, p4 - p1), "")
                End If
                s = Mid(s, p2 + 1)
            Else
                Exit Do
            End If
        Loop
        s = ss + s
    Wend
    
    RemoveHTMLCode = s
    
    Exit Function
    
AbortFcn:

    RemoveHTMLCode = sCode

End Function

