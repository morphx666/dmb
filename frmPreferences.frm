VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{2A2AD7CA-AC77-46F3-84DC-115021432312}#1.0#0"; "HREF.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmPreferences 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preferences"
   ClientHeight    =   9075
   ClientLeft      =   1965
   ClientTop       =   3780
   ClientWidth     =   12225
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreferences.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9075
   ScaleWidth      =   12225
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picLinks 
      Height          =   3285
      Left            =   765
      ScaleHeight     =   3225
      ScaleWidth      =   5250
      TabIndex        =   59
      Top             =   7020
      Width           =   5310
      Begin VB.CheckBox chkVL_External 
         Caption         =   "Verify External Links"
         Height          =   195
         Left            =   0
         TabIndex        =   65
         Top             =   1425
         Width           =   4335
      End
      Begin xfxLine3D.ucLine3D uc3DLine3 
         Height          =   30
         Left            =   30
         TabIndex        =   64
         Top             =   1200
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   53
      End
      Begin VB.CheckBox chkVL_Compiling 
         Caption         =   "Compiling"
         Height          =   210
         Left            =   285
         TabIndex        =   63
         Top             =   795
         Width           =   3795
      End
      Begin VB.CheckBox chkVL_Saving 
         Caption         =   "Saving"
         Height          =   210
         Left            =   285
         TabIndex        =   62
         Top             =   555
         Width           =   3795
      End
      Begin VB.CheckBox chkVL_Opening 
         Caption         =   "Opening"
         Height          =   210
         Left            =   285
         TabIndex        =   61
         Top             =   315
         Width           =   3795
      End
      Begin VB.Label lblVerifyLinks 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verify Links When"
         Height          =   195
         Left            =   0
         TabIndex        =   60
         Top             =   45
         Width           =   1275
      End
   End
   Begin VB.PictureBox picDisplay 
      Height          =   4185
      Left            =   6345
      ScaleHeight     =   4125
      ScaleWidth      =   5250
      TabIndex        =   35
      Top             =   4200
      Width           =   5310
      Begin VB.CommandButton cmdSetDefaultStyles 
         Caption         =   "Defaults"
         Height          =   315
         Left            =   4395
         TabIndex        =   54
         Top             =   3750
         Width           =   780
      End
      Begin VB.Frame frameSIS 
         Caption         =   "Special Items Styles"
         Height          =   1635
         Left            =   90
         TabIndex        =   49
         Top             =   2040
         Width           =   5085
         Begin VB.CommandButton cmdColor 
            Caption         =   "Color"
            Height          =   255
            Index           =   6
            Left            =   195
            TabIndex        =   67
            Top             =   1215
            Width           =   690
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "Color"
            Height          =   255
            Index           =   5
            Left            =   2715
            TabIndex        =   53
            Top             =   570
            Width           =   690
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "Color"
            Height          =   255
            Index           =   4
            Left            =   180
            TabIndex        =   51
            Top             =   570
            Width           =   690
         End
         Begin VB.Image Image7 
            Height          =   240
            Left            =   195
            Picture         =   "frmPreferences.frx":038A
            Top             =   960
            Width           =   240
         End
         Begin VB.Label lblNoCompile 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Items excluded from compilation"
            Height          =   195
            Left            =   495
            TabIndex        =   68
            Top             =   975
            Width           =   2310
         End
         Begin VB.Label lblBrokenLinks 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Items with broken links"
            Height          =   195
            Left            =   3015
            TabIndex        =   52
            Top             =   330
            Width           =   1635
         End
         Begin VB.Image Image6 
            Height          =   240
            Left            =   2715
            Picture         =   "frmPreferences.frx":0914
            Top             =   315
            Width           =   240
         End
         Begin VB.Label lblDisabled 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Disabled Items"
            Height          =   195
            Left            =   480
            TabIndex        =   50
            Top             =   330
            Width           =   1050
         End
         Begin VB.Image Image5 
            Height          =   240
            Left            =   195
            Picture         =   "frmPreferences.frx":0E9E
            Top             =   360
            Width           =   240
         End
      End
      Begin VB.Frame frameMIS 
         Caption         =   "Menu Items Style"
         Height          =   1950
         Left            =   90
         TabIndex        =   36
         Top             =   30
         Width           =   5085
         Begin VB.CheckBox chkBold 
            Caption         =   "Bold"
            Height          =   255
            Index           =   3
            Left            =   2715
            Style           =   1  'Graphical
            TabIndex        =   58
            Top             =   1410
            Width           =   690
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "Bold"
            Height          =   255
            Index           =   2
            Left            =   2715
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   630
            Width           =   690
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "Bold"
            Height          =   255
            Index           =   1
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   1410
            Width           =   690
         End
         Begin VB.CheckBox chkBold 
            Caption         =   "Bold"
            Height          =   255
            Index           =   0
            Left            =   180
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   630
            Width           =   690
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "Color"
            Height          =   255
            Index           =   3
            Left            =   3540
            TabIndex        =   48
            Top             =   1410
            Width           =   690
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Font"
            Height          =   255
            Index           =   3
            Left            =   2715
            TabIndex        =   47
            Top             =   1410
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "Color"
            Height          =   255
            Index           =   2
            Left            =   3540
            TabIndex        =   45
            Top             =   630
            Width           =   690
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Font"
            Height          =   255
            Index           =   2
            Left            =   2715
            TabIndex        =   44
            Top             =   630
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "Color"
            Height          =   255
            Index           =   1
            Left            =   1005
            TabIndex        =   42
            Top             =   1410
            Width           =   690
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Font"
            Height          =   255
            Index           =   1
            Left            =   180
            TabIndex        =   41
            Top             =   1410
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.CommandButton cmdColor 
            Caption         =   "Color"
            Height          =   255
            Index           =   0
            Left            =   1005
            TabIndex        =   39
            Top             =   630
            Width           =   690
         End
         Begin VB.CommandButton cmdFont 
            Caption         =   "Font"
            Height          =   255
            Index           =   0
            Left            =   180
            TabIndex        =   38
            Top             =   630
            Visible         =   0   'False
            Width           =   690
         End
         Begin VB.Label lblCommands 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Commands"
            Height          =   195
            Left            =   3015
            TabIndex        =   46
            Top             =   1170
            Width           =   780
         End
         Begin VB.Image Image4 
            Height          =   240
            Left            =   2715
            Picture         =   "frmPreferences.frx":1428
            Top             =   1215
            Width           =   240
         End
         Begin VB.Label lblGroups 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Groups"
            Height          =   195
            Left            =   3015
            TabIndex        =   43
            Top             =   390
            Width           =   510
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   2715
            Picture         =   "frmPreferences.frx":19B2
            Top             =   435
            Width           =   240
         End
         Begin VB.Label lblToolbarItems 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Toolbar Items"
            Height          =   195
            Left            =   480
            TabIndex        =   40
            Top             =   1170
            Width           =   990
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   180
            Picture         =   "frmPreferences.frx":1F3C
            Top             =   1215
            Width           =   240
         End
         Begin VB.Label lblToolbars 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Toolbars"
            Height          =   195
            Left            =   480
            TabIndex        =   37
            Top             =   390
            Width           =   615
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   195
            Picture         =   "frmPreferences.frx":2086
            Top             =   375
            Width           =   240
         End
      End
      Begin VB.Label lblShadow 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "lblshadow"
         Height          =   195
         Index           =   0
         Left            =   1740
         TabIndex        =   69
         Top             =   3795
         Visible         =   0   'False
         Width           =   705
      End
   End
   Begin VB.PictureBox picGeneral 
      Height          =   4905
      Left            =   240
      ScaleHeight     =   4845
      ScaleWidth      =   5250
      TabIndex        =   0
      Top             =   585
      Width           =   5310
      Begin href.uchref1 uchrefNET35 
         Height          =   315
         Left            =   1815
         TabIndex        =   72
         Top             =   2610
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         Caption         =   "(Requires .NET 3.5)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "Tahoma"
         FontSize        =   8.25
         ForeColor       =   16711680
         URL             =   "http://www.microsoft.com/downloads/details.aspx?FamilyID=333325FD-AE52-4E35-B531-508D977D32A6&displaylang=en"
      End
      Begin VB.CheckBox chkUnicodeInput 
         Caption         =   "Enable Unicode Input"
         Height          =   210
         Left            =   0
         TabIndex        =   71
         Top             =   2670
         Width           =   1905
      End
      Begin VB.CheckBox chkUseEasyActions 
         Caption         =   "Use Simplified Events Selection Interface"
         Height          =   210
         Left            =   0
         TabIndex        =   70
         Top             =   2400
         Width           =   5175
      End
      Begin VB.CheckBox chkCleanPreview 
         Caption         =   "Show Project Information in the Preview window"
         Height          =   210
         Left            =   0
         TabIndex        =   66
         Top             =   2130
         Width           =   5175
      End
      Begin VB.CommandButton cmdDefaultCP 
         Caption         =   "Default"
         Height          =   315
         Left            =   2805
         TabIndex        =   15
         Top             =   3930
         Width           =   780
      End
      Begin VB.ComboBox cmbCharset 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   3930
         Width           =   2745
      End
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   0
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2970
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   53
      End
      Begin VB.CheckBox chkDefaultToMapView 
         Caption         =   "Default to Map View"
         Height          =   210
         Left            =   0
         TabIndex        =   8
         Top             =   1860
         Width           =   5175
      End
      Begin VB.CheckBox chkLock 
         Caption         =   "Enable Project Locking"
         Height          =   210
         Left            =   0
         TabIndex        =   7
         Top             =   1590
         Width           =   5175
      End
      Begin VB.CheckBox chkAdvancedLC 
         Caption         =   "Use Advanced Menus Installation Tools"
         Height          =   210
         Left            =   0
         TabIndex        =   6
         Top             =   1325
         Width           =   5175
      End
      Begin VB.CommandButton cmdResetTips 
         Caption         =   "Reset Tips"
         Height          =   345
         Left            =   0
         TabIndex        =   17
         Top             =   4470
         Width           =   1305
      End
      Begin VB.CheckBox chkShowPP 
         Caption         =   "Show Project Properties on New Project"
         Height          =   210
         Left            =   0
         TabIndex        =   5
         Top             =   1060
         Width           =   5175
      End
      Begin SmartButtonProject.SmartButton cmdInfo 
         Height          =   360
         Left            =   2805
         TabIndex        =   12
         Top             =   3225
         Width           =   420
         _ExtentX        =   741
         _ExtentY        =   635
         Picture         =   "frmPreferences.frx":21D0
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
      Begin VB.CheckBox chkOpenLastProject 
         Caption         =   "Open last project when DMB starts"
         Height          =   210
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   5175
      End
      Begin VB.CheckBox chkAutoRecover 
         Caption         =   "Auto Recover"
         Height          =   210
         Left            =   0
         TabIndex        =   2
         Top             =   265
         Width           =   5175
      End
      Begin VB.CheckBox chkShowNag 
         Caption         =   "Show logo splash-screen when DMB starts"
         Height          =   210
         Left            =   0
         TabIndex        =   3
         Top             =   530
         Width           =   5175
      End
      Begin VB.CheckBox chkDisUR 
         Caption         =   "Enable Undo/Redo"
         Height          =   210
         Left            =   0
         TabIndex        =   4
         Top             =   795
         Width           =   5175
      End
      Begin VB.ComboBox cmbLanguages 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   3255
         Width           =   2745
      End
      Begin xfxLine3D.ucLine3D uc3DLine2 
         Height          =   30
         Left            =   0
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   4335
         Width           =   5235
         _ExtentX        =   9234
         _ExtentY        =   53
      End
      Begin VB.Label lblCodepage 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Character Encoding for Preview"
         Height          =   195
         Left            =   0
         TabIndex        =   13
         Top             =   3705
         Width           =   2850
      End
      Begin VB.Label lblLng 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Language"
         Height          =   195
         Left            =   0
         TabIndex        =   10
         Top             =   3015
         Width           =   705
      End
   End
   Begin MSComctlLib.TabStrip tsOptions 
      Height          =   5475
      Left            =   60
      TabIndex        =   34
      Top             =   105
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   9657
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
            Caption         =   "Formatting"
            Key             =   "tsFormatting"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Display"
            Key             =   "tsDisplay"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Links Verification"
            Key             =   "tsLinks"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFormatting 
      Height          =   3285
      Left            =   6075
      ScaleHeight     =   3225
      ScaleWidth      =   5250
      TabIndex        =   18
      Top             =   765
      Width           =   5310
      Begin VB.TextBox txtSepH 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   165
         TabIndex        =   20
         Text            =   "123"
         Top             =   270
         Width           =   360
      End
      Begin VB.Frame frameGrp 
         Caption         =   "Groups inherit their style from"
         Height          =   810
         Left            =   0
         TabIndex        =   29
         Top             =   2235
         Width           =   4575
         Begin VB.OptionButton opGrpInherit 
            Caption         =   "Program Defaults"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   30
            Top             =   270
            Value           =   -1  'True
            Width           =   4125
         End
         Begin VB.OptionButton opGrpInherit 
            Caption         =   "First Group on the Project"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   31
            Top             =   495
            Width           =   4125
         End
      End
      Begin VB.Frame frameCmd 
         Caption         =   "Commands inherit their style from"
         Height          =   810
         Left            =   0
         TabIndex        =   26
         Top             =   1350
         Width           =   4575
         Begin VB.OptionButton opCmdInherit 
            Caption         =   "Parent Group"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   27
            Top             =   270
            Value           =   -1  'True
            Width           =   4170
         End
         Begin VB.OptionButton opCmdInherit 
            Caption         =   "First Command on the same Group"
            Height          =   195
            Index           =   1
            Left            =   210
            TabIndex        =   28
            Top             =   495
            Width           =   4170
         End
      End
      Begin VB.TextBox txtImageSpace 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   165
         TabIndex        =   23
         Text            =   "123"
         Top             =   915
         Width           =   360
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   285
         Left            =   525
         TabIndex        =   24
         Top             =   915
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtImageSpace"
         BuddyDispid     =   196661
         OrigLeft        =   525
         OrigTop         =   885
         OrigRight       =   720
         OrigBottom      =   1170
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   285
         Left            =   525
         TabIndex        =   21
         Top             =   270
         Width           =   195
         _ExtentX        =   344
         _ExtentY        =   503
         _Version        =   393216
         Value           =   3
         BuddyControl    =   "txtSepH"
         BuddyDispid     =   196656
         OrigLeft        =   525
         OrigTop         =   255
         OrigRight       =   720
         OrigBottom      =   540
         Max             =   30
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblISCaption 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         Height          =   195
         Left            =   1260
         TabIndex        =   25
         Top             =   945
         Width           =   555
      End
      Begin VB.Image imgPic 
         Height          =   240
         Left            =   975
         Picture         =   "frmPreferences.frx":256A
         Top             =   930
         Width           =   240
      End
      Begin VB.Shape shpCmd 
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   915
         Top             =   900
         Width           =   1215
      End
      Begin VB.Line lnSep 
         X1              =   1005
         X2              =   2030
         Y1              =   390
         Y2              =   390
      End
      Begin VB.Label lblSepH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Separator Height"
         Height          =   195
         Left            =   0
         TabIndex        =   19
         Top             =   30
         Width           =   1230
      End
      Begin VB.Label lblImgH 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Images Space"
         Height          =   195
         Left            =   0
         TabIndex        =   22
         Top             =   675
         Width           =   1005
      End
      Begin VB.Shape shpSepSpace 
         BackStyle       =   1  'Opaque
         Height          =   150
         Left            =   915
         Top             =   330
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3780
      TabIndex        =   32
      Top             =   5640
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4875
      TabIndex        =   33
      Top             =   5640
      Width           =   900
   End
End
Attribute VB_Name = "frmPreferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Translators() As String
Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, col As Long) As Long

Private Sub chkBold_Click(Index As Integer)

    On Error Resume Next

    picDisplay.SetFocus

    Select Case Index
        Case 0: lblToolbars.FontBold = chkBold(Index).Value
        Case 1: lblToolbarItems.FontBold = chkBold(Index).Value
        Case 2: lblGroups.FontBold = chkBold(Index).Value
        Case 3: lblCommands.FontBold = chkBold(Index).Value
    End Select

End Sub

Private Sub chkUnicodeInput_Click()

    Static isRecall As Boolean
    
    If isRecall Then Exit Sub
    isRecall = True

    If uchrefNET35.Visible Then
        chkUnicodeInput.Value = vbUnchecked
        If MsgBox("In order to use the Unicode Input tool you will need to install Microsoft's .NET 3.5 Framework" + vbCrLf + "Do you want to download it now?", vbQuestion Or vbYesNo, "Missing .NET 3.5") = vbYes Then
            uchrefNET35_Click
        End If
    End If
    
    isRecall = False

End Sub

Private Sub chkVL_Compiling_Click()

    Update_chkVL_External_Status

End Sub

Private Sub chkVL_External_Click()

    Update_chkVL_External_Status

End Sub

Private Sub chkVL_Opening_Click()

    Update_chkVL_External_Status

End Sub

Private Sub Update_chkVL_External_Status()

    #If LITE = 1 Then
        Dim ShowMsg As Boolean
        
        ShowMsg = CBool(chkVL_Compiling.Value Or chkVL_Saving.Value Or chkVL_Opening.Value)
        
        chkVL_Compiling.Value = vbUnchecked
        chkVL_Saving.Value = vbUnchecked
        chkVL_Opening.Value = vbUnchecked
        chkVL_External.Value = vbUnchecked
        chkVL_External.Enabled = False
        
        If ShowMsg Then frmMain.ShowLITELImitationInfo 2
    #Else
        chkVL_External.Enabled = CBool(chkVL_Compiling.Value Or chkVL_Saving.Value Or chkVL_Opening.Value)
    #End If

End Sub

Private Sub chkVL_Saving_Click()

    Update_chkVL_External_Status

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdColor_Click(Index As Integer)

    picDisplay.SetFocus

    With cmdColor(Index)
        Select Case Index
            Case 0: TranslateColor lblToolbars.ForeColor, 0, SelColor
            Case 1: TranslateColor lblToolbarItems.ForeColor, 0, SelColor
            Case 2: TranslateColor lblGroups.ForeColor, 0, SelColor
            Case 3: TranslateColor lblCommands.ForeColor, 0, SelColor
            Case 4: TranslateColor lblDisabled.ForeColor, 0, SelColor
            Case 5: TranslateColor lblBrokenLinks.ForeColor, 0, SelColor
            Case 6: TranslateColor lblNoCompile.ForeColor, 0, SelColor
        End Select
        SelColor_CanBeTransparent = False
        frmColorPicker.Show vbModal
        If SelColor <> -1 Then
            Select Case Index
                Case 0: lblToolbars.ForeColor = SelColor
                Case 1: lblToolbarItems.ForeColor = SelColor
                Case 2: lblGroups.ForeColor = SelColor
                Case 3: lblCommands.ForeColor = SelColor
                Case 4: lblDisabled.ForeColor = SelColor
                Case 5: lblBrokenLinks.ForeColor = SelColor
                Case 6: lblNoCompile.ForeColor = SelColor
            End Select
        End If
    End With

End Sub

Private Sub cmdDefaultCP_Click()

    Dim i As Integer
    Dim cs() As CodePagesDef
    Dim DefCP As String
    
    DefCP = "1252"
    If LenB(QueryValue(HKEY_CLASSES_ROOT, "MIME\Database\Codepage\50001", "BodyCharset")) <> 0 Then
        DefCP = "50001"
    End If
    
    cs = GetSysCharsets
    For i = 1 To UBound(cs)
        If cs(i).CodePage = DefCP Then
            cmbCharset.ListIndex = i - 1
            Exit For
        End If
    Next i
    
    cmbCharset.SetFocus

End Sub

'Private Sub cmdFont_Click(Index As Integer)
'
'    Dim lbl As Label
'
'    Me.SetFocus
'
'    Select Case Index
'        Case 0: Set lbl = lblToolbars
'        Case 1: Set lbl = lblToolbarItems
'        Case 2: Set lbl = lblGroups
'        Case 3: Set lbl = lblCommands
'    End Select
'
'    SelFont.Bold = lbl.Font.Bold
'    SelFont.Italic = lbl.Font.Italic
'    SelFont.Name = lbl.Font.Name
'    SelFont.Size = pt2px(lbl.Font.Size)
'    SelFont.Underline = lbl.Font.Underline
'
'    With frmFontDialog
'        .Show vbModal
'        CenterForm frmFontDialog
'        Do
'            DoEvents
'        Loop While .Visible
'    End With
'
'    If SelFont.IsValid Then
'        With lbl.Font
'            .Name = SelFont.Name
'            .Size = px2pt(SelFont.Size)
'            .Bold = SelFont.Bold
'            .Italic = SelFont.Italic
'            .Underline = SelFont.Underline
'        End With
'    End If
'
'End Sub

Private Sub cmdInfo_Click()

    MsgBox GetLocalizedStr(427) + " " + vbCrLf + Translators(cmbLanguages.ListIndex), vbInformation + vbOKOnly, "Language Package Information"
    cmbLanguages.SetFocus

End Sub

Private Sub cmdOK_Click()

    Dim cs() As CodePagesDef

    If IsDEMO Then
        Preferences.ShowNag = True
    Else
        Preferences.ShowNag = -chkShowNag.Value
    End If
    
    With Preferences
        .OpenLastProject = -chkOpenLastProject.Value
        .AutoRecover = -chkAutoRecover.Value
        .SepHeight = Val(txtSepH.Text)
        .ImgSpace = Val(txtImageSpace.Text)
        
        .CommandsInheritance = IIf(opCmdInherit(0).Value, icDefault, icFirst)
        .GroupsInheritance = IIf(opGrpInherit(0).Value, icDefault, icFirst)
        
        .EnableUndoRedo = -chkDisUR.Value
        .ShowPPOnNewProject = -chkShowPP.Value
        
        .language = Mid(cmbLanguages.Text, 2, InStr(cmbLanguages.Text, ")") - 2)
        
        .UseInstallMenus = Not -chkAdvancedLC.Value
        .UseMapView = -chkDefaultToMapView.Value
        .LockProjects = -chkLock.Value
        .UseEasyActions = -chkUseEasyActions.Value
        
        cs = GetSysCharsets
        .CodePage = cs(cmbCharset.ListIndex + 1).CodePage
        
        .ToolbarStyle.Color = lblToolbars.ForeColor
        .ToolbarStyle.Font.FontBold = lblToolbars.FontBold
        'Font2tFont lblToolbars.Font, .ToolbarStyle.Font
        
        .ToolbarItemStyle.Color = lblToolbarItems.ForeColor
        .ToolbarItemStyle.Font.FontBold = lblToolbarItems.FontBold
        'Font2tFont lblToolbarItems.Font, .ToolbarItemStyle.Font
        
        .GroupStyle.Color = lblGroups.ForeColor
        .GroupStyle.Font.FontBold = lblGroups.FontBold
        'Font2tFont lblGroups.Font, .GroupStyle.Font
        
        .CommandStyle.Color = lblCommands.ForeColor
        .CommandStyle.Font.FontBold = lblCommands.FontBold
        'Font2tFont lblCommands.Font, .CommandStyle.Font
        
        .DisabledItem = lblDisabled.ForeColor
        .BrokenLink = lblBrokenLinks.ForeColor
        .NoCompileItem = lblNoCompile.ForeColor
        
        With .VerifyLinksOptions
            .VerifyExternalLinks = -chkVL_External.Value
            .VerifyOptions = chkVL_Compiling.Value * lvcVerifyWhenCompiling + _
                            chkVL_Opening.Value * lvcVerifyWhenOpening + _
                            chkVL_Saving.Value * lvcVerifyWhenSaving
        End With
        
        .ShowCleanPreview = Not -chkCleanPreview.Value
        
        .EnableUnicodeInput = -chkUnicodeInput.Value
    End With
    
    If chkUnicodeInput.Value = vbChecked Then LoadUnicodeTool
    
    Unload Me

End Sub

Private Sub cmdResetTips_Click()

    On Error Resume Next
    DeleteSetting App.EXEName, "TipsSystem"

End Sub

Private Sub cmdSetDefaultStyles_Click()

    lblToolbars.ForeColor = &H80000012
    chkBold(0).Value = vbChecked
    
    lblToolbarItems.ForeColor = &H80000012
    chkBold(1).Value = vbChecked
    
    lblGroups.ForeColor = &H80000012
    chkBold(2).Value = vbChecked
    
    lblCommands.ForeColor = &H80000012
    chkBold(3).Value = vbUnchecked
    
    lblDisabled.ForeColor = &H80000011
    lblBrokenLinks.ForeColor = vbRed
    lblNoCompile.ForeColor = &H80000013

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        Select Case tsOptions.SelectedItem.key
            Case "tsGeneral"
                showHelp "dialogs/preferences.htm"
            Case "tsFormatting"
                showHelp "dialogs/preferences_formatting.htm"
            Case "tsDisplay"
                showHelp "dialogs/preferences_display.htm"
            Case "tsLinks"
                showHelp "dialogs/preferences_links.htm"
        End Select
    End If

End Sub

Private Sub Form_Load()

    Dim ctrl As Control
    
    Width = 5940
    Height = cmdOK.Top + cmdOK.Height + 4 * Screen.TwipsPerPixelY + GetClientTop(Me.hwnd)
    
    CenterForm Me
    LocalizeUI

    If IsDEMO Then
        chkShowNag.Enabled = False
        chkShowNag.Value = vbChecked
    Else
        chkShowNag.Value = Abs(Preferences.ShowNag)
    End If
    
    For Each ctrl In Controls
        If TypeOf ctrl Is PictureBox Then
            If LenB(ctrl.tag) = 0 Then
                ctrl.BorderStyle = 0
                ctrl.Move picGeneral.Left, picGeneral.Top, picGeneral.Width, picGeneral.Height
            End If
        End If
    Next ctrl
    picGeneral.ZOrder 0
    
    With Preferences
        chkOpenLastProject.Value = Abs(.OpenLastProject)
        chkAutoRecover.Value = Abs(.AutoRecover)
        txtSepH.Text = CStr(.SepHeight)
        txtImageSpace.Text = CStr(.ImgSpace)
        
        opCmdInherit(0).Value = (.CommandsInheritance = icDefault)
        opCmdInherit(1).Value = (.CommandsInheritance = icFirst)
        
        opGrpInherit(0).Value = (.GroupsInheritance = icDefault)
        opGrpInherit(1).Value = (.GroupsInheritance = icFirst)
        
        chkDisUR.Value = Abs(.EnableUndoRedo)
        chkShowPP.Value = Abs(.ShowPPOnNewProject)
        
        chkAdvancedLC.Value = Abs(Not .UseInstallMenus)
        chkDefaultToMapView.Value = Abs(.UseMapView)
        chkLock.Value = Abs(.LockProjects)
        chkUseEasyActions.Value = Abs(.UseEasyActions)
        
        lblToolbars.ForeColor = Preferences.ToolbarStyle.Color
        chkBold(0).Value = Abs(Preferences.ToolbarStyle.Font.FontBold)
        'tFont2Font Preferences.ToolbarStyle.Font, lblToolbars.Font
        
        lblToolbarItems.ForeColor = Preferences.ToolbarItemStyle.Color
        chkBold(1).Value = Abs(Preferences.ToolbarItemStyle.Font.FontBold)
        'tFont2Font Preferences.ToolbarItemStyle.Font, lblToolbarItems.Font
        
        lblGroups.ForeColor = Preferences.GroupStyle.Color
        chkBold(2).Value = Abs(Preferences.GroupStyle.Font.FontBold)
        'tFont2Font Preferences.GroupStyle.Font, lblGroups.Font
        
        lblCommands.ForeColor = Preferences.CommandStyle.Color
        chkBold(3).Value = Abs(Preferences.CommandStyle.Font.FontBold)
        'tFont2Font Preferences.CommandStyle.Font, lblCommands.Font
        
        lblDisabled.ForeColor = Preferences.DisabledItem
        lblBrokenLinks.ForeColor = Preferences.BrokenLink
        lblNoCompile.ForeColor = Preferences.NoCompileItem
        
        With .VerifyLinksOptions
            chkVL_Compiling.Value = Abs((.VerifyOptions And lvcVerifyWhenCompiling) = lvcVerifyWhenCompiling)
            chkVL_Opening.Value = Abs((.VerifyOptions And lvcVerifyWhenOpening) = lvcVerifyWhenOpening)
            chkVL_Saving.Value = Abs((.VerifyOptions And lvcVerifyWhenSaving) = lvcVerifyWhenSaving)
            chkVL_External.Value = Abs(.VerifyExternalLinks)
        End With
        
        chkCleanPreview.Value = Abs(Not .ShowCleanPreview)
        
        chkUnicodeInput.Value = Abs(.EnableUnicodeInput)
    End With
    
    FillLanguagesCombo
    FillCharsetsCombo
    
    FixCtrls4Skin Me
    
    BuildUsedColorsArray
    
    Update_chkVL_External_Status
    
    '----------------
    
    createShadow lblToolbars
    createShadow lblGroups
    createShadow lblToolbarItems
    createShadow lblCommands
    createShadow lblDisabled
    createShadow lblBrokenLinks
    createShadow lblNoCompile
    
    uchrefNET35.Visible = Not FolderExists(GetWindowsDir + "\Microsoft.NET\Framework\v3.5")
    
End Sub

Private Sub createShadow(lbl As Label)

    Load lblShadow(lblShadow.Count)
    With lblShadow(lblShadow.Count - 1)
        .caption = lbl.caption
        .Move lbl.Left + 15, lbl.Top + 15
        .ForeColor = RGB(128, 128, 128)
        Set .Font = lbl.Font
        Set .Container = lbl.Container
        .ZOrder 1
        .Visible = True
    End With

End Sub

Private Sub FillCharsetsCombo()

    Dim i As Integer
    Dim cs() As CodePagesDef
    
    cs = GetSysCharsets
    For i = 1 To UBound(cs)
        cmbCharset.AddItem cs(i).Description
        If cs(i).CodePage = Preferences.CodePage Then
            cmbCharset.ListIndex = cmbCharset.NewIndex
        End If
    Next i

End Sub

Private Sub FillLanguagesCombo()

    Dim fn As String
    Dim ff As Integer
    Dim lname As String
    Dim i As Integer
    
    ReDim Translators(0)
    
    ff = FreeFile
    fn = Dir(AppPath + "lang\*")
    While LenB(fn) <> 0
        If InStr(fn, ".") = 0 Then
            Open AppPath + "lang\" + fn For Input As #ff
                Line Input #ff, lname
                cmbLanguages.AddItem "(" + fn + ") " + lname
                Line Input #ff, lname
                Translators(UBound(Translators)) = lname
                ReDim Preserve Translators(UBound(Translators) + 1)
            Close #ff
        End If
        fn = Dir
    Wend
    
    cmbLanguages.ListIndex = 0
    For i = 0 To cmbLanguages.ListCount - 1
        If Mid(cmbLanguages.List(i), 2, InStr(cmbLanguages.List(i), ")") - 2) = Preferences.language Then
            cmbLanguages.ListIndex = i
            Exit For
        End If
    Next i

End Sub

Private Sub tsOptions_Click()

    Select Case tsOptions.SelectedItem.key
        Case "tsGeneral"
            picGeneral.ZOrder 0
        Case "tsFormatting"
            picFormatting.ZOrder 0
        Case "tsDisplay"
            picDisplay.ZOrder 0
        Case "tsLinks"
            picLinks.ZOrder 0
    End Select

End Sub

Private Sub txtImageSpace_Change()

    UpdateSamples

End Sub

Private Sub txtImageSpace_GotFocus()

    SelAll txtImageSpace

End Sub

Private Sub txtImageSpace_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtSepH_Change()

    UpdateSamples

End Sub

Private Sub txtSepH_GotFocus()

    SelAll txtSepH

End Sub

Private Sub txtSepH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub LocalizeUI()

    caption = GetLocalizedStr(450)
    
    tsOptions.Tabs(1).caption = GetLocalizedStr(434)
    chkOpenLastProject.caption = GetLocalizedStr(435)
    chkAutoRecover.caption = GetLocalizedStr(436)
    chkShowNag.caption = GetLocalizedStr(437)
    chkDisUR.caption = GetLocalizedStr(438)
    chkShowPP.caption = GetLocalizedStr(679)
    chkAdvancedLC.caption = GetLocalizedStr(755)
    chkLock.caption = GetLocalizedStr(756)
    chkDefaultToMapView.caption = GetLocalizedStr(757)
    
    lblLng.caption = GetLocalizedStr(440)
    
    lblCodepage.caption = GetLocalizedStr(919)
    cmdDefaultCP.caption = GetLocalizedStr(400)
    
    tsOptions.Tabs(2).caption = GetLocalizedStr(441)
    lblSepH.caption = GetLocalizedStr(442)
    lblImgH.caption = GetLocalizedStr(443)
    frameCmd.caption = GetLocalizedStr(444)
    opCmdInherit(0).caption = GetLocalizedStr(445)
    opCmdInherit(1).caption = GetLocalizedStr(446)
    frameGrp.caption = GetLocalizedStr(447)
    opGrpInherit(0).caption = GetLocalizedStr(448)
    opGrpInherit(1).caption = GetLocalizedStr(449)
    
    tsOptions.Tabs(3).caption = GetLocalizedStr(952)
    lblVerifyLinks.caption = GetLocalizedStr(947)
    chkVL_Opening.caption = GetLocalizedStr(948)
    chkVL_Saving.caption = GetLocalizedStr(949)
    chkVL_Compiling.caption = GetLocalizedStr(950)
    chkVL_External.caption = GetLocalizedStr(951)
    
    tsOptions.Tabs(4).caption = GetLocalizedStr(953)
    frameMIS.caption = GetLocalizedStr(939)
    lblToolbars.caption = GetLocalizedStr(940)
    lblGroups.caption = GetLocalizedStr(941)
    lblToolbarItems.caption = GetLocalizedStr(942)
    lblCommands.caption = GetLocalizedStr(943)
    frameSIS.caption = GetLocalizedStr(944)
    chkBold(0).caption = GetLocalizedStr(510)
    chkBold(1).caption = GetLocalizedStr(510)
    chkBold(2).caption = GetLocalizedStr(510)
    chkBold(3).caption = GetLocalizedStr(510)
    cmdColor(0).caption = GetLocalizedStr(212)
    cmdColor(1).caption = GetLocalizedStr(212)
    cmdColor(2).caption = GetLocalizedStr(212)
    cmdColor(3).caption = GetLocalizedStr(212)
    cmdColor(4).caption = GetLocalizedStr(212)
    cmdColor(5).caption = GetLocalizedStr(212)
    cmdSetDefaultStyles.caption = GetLocalizedStr(954)
    lblNoCompile.caption = GetLocalizedStr(986)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
        cmdDefaultCP.Width = SetCtrlWidth(cmdDefaultCP)
        cmdSetDefaultStyles.Width = SetCtrlWidth(cmdSetDefaultStyles)
        cmdSetDefaultStyles.Left = frameSIS.Left + frameSIS.Width - cmdSetDefaultStyles.Width
    End If
    
    FixContolsWidth Me
    
End Sub

Private Sub UpdateSamples()

    With shpSepSpace
        .Height = Val(txtSepH.Text) * 15
        .Height = CInt(.Height / 2) * 2
        .Move .Left, txtSepH.Top + (txtSepH.Height - .Height) / 2 - 15
    End With
    
    lblISCaption.Left = imgPic.Left + imgPic.Width + Val(txtImageSpace.Text) * 15
    
End Sub

Private Sub uchrefNET35_Click()

    RunShellExecute "Open", uchrefNET35.url, 0, 0, 0

End Sub
