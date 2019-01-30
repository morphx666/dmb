VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmExportHTML 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Export As HTML"
   ClientHeight    =   8100
   ClientLeft      =   3615
   ClientTop       =   4275
   ClientWidth     =   13980
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8100
   ScaleWidth      =   13980
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmCodeGen 
      Caption         =   "Code Compliance"
      Height          =   1050
      Left            =   195
      TabIndex        =   64
      Top             =   6345
      Width           =   6540
      Begin VB.OptionButton opCodeCompliance 
         Caption         =   "XHTML 1.0 Transitional"
         Height          =   195
         Index           =   1
         Left            =   525
         TabIndex        =   66
         Top             =   645
         Width           =   4050
      End
      Begin VB.OptionButton opCodeCompliance 
         Caption         =   "HTML 4.01 Transitional"
         Height          =   195
         Index           =   0
         Left            =   525
         TabIndex        =   65
         Top             =   360
         Width           =   4050
      End
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   8385
      Top             =   6915
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "0"
      DialogTitle     =   "Select Project"
      Filter          =   "DHTML Menu Builder Projects|*.dmb"
      MaxFileSize     =   1024
   End
   Begin VB.PictureBox picColTree 
      Height          =   5850
      Left            =   7080
      ScaleHeight     =   5790
      ScaleWidth      =   6525
      TabIndex        =   30
      Top             =   465
      Width           =   6585
      Begin VB.CommandButton cmdDefaults 
         Caption         =   "Defaults"
         Height          =   330
         Left            =   5520
         TabIndex        =   57
         Top             =   5430
         Width           =   960
      End
      Begin VB.Frame Frame5 
         Caption         =   "Options"
         Height          =   5025
         Left            =   2865
         TabIndex        =   54
         Top             =   315
         Width           =   3660
         Begin VB.ComboBox cmbExpColAllPlacement 
            Height          =   315
            ItemData        =   "frmExportHTML.frx":0000
            Left            =   1155
            List            =   "frmExportHTML.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   62
            Top             =   2325
            Width           =   1635
         End
         Begin VB.TextBox txtColAll 
            Height          =   285
            Left            =   450
            TabIndex        =   61
            Text            =   "Collapse All"
            Top             =   1950
            Width           =   2340
         End
         Begin VB.TextBox txtExpAll 
            Height          =   285
            Left            =   450
            TabIndex        =   60
            Text            =   "Expand All"
            Top             =   1590
            Width           =   2340
         End
         Begin VB.CheckBox chkColExpAll 
            Caption         =   "Include Expand/Collapse All options"
            Height          =   225
            Left            =   165
            TabIndex        =   59
            Top             =   1305
            Width           =   3300
         End
         Begin VB.CheckBox chkSingleSel 
            Caption         =   "Single Select"
            Height          =   225
            Left            =   165
            TabIndex        =   56
            Top             =   690
            Width           =   3300
         End
         Begin VB.CheckBox chkExpItemsLinks 
            Caption         =   "Expandable Items can have Links"
            Height          =   225
            Left            =   165
            TabIndex        =   55
            Top             =   345
            Value           =   1  'Checked
            Width           =   3300
         End
         Begin xfxLine3D.ucLine3D uc3DLine4 
            Height          =   30
            Left            =   75
            TabIndex        =   58
            Top             =   1080
            Width           =   3495
            _ExtentX        =   6165
            _ExtentY        =   53
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Place at"
            Height          =   195
            Left            =   450
            TabIndex        =   63
            Top             =   2385
            Width           =   570
         End
      End
      Begin VB.CheckBox chkTree 
         Caption         =   "Create Collapsible Tree"
         Height          =   255
         Left            =   0
         TabIndex        =   53
         Top             =   0
         Width           =   4260
      End
      Begin VB.Frame Frame3 
         Caption         =   "Images"
         Height          =   5025
         Left            =   0
         TabIndex        =   31
         Top             =   315
         Width           =   2775
         Begin VB.TextBox txtLeftH 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1755
            TabIndex        =   37
            Text            =   "16"
            Top             =   3525
            WhatsThisHelpID =   20340
            Width           =   420
         End
         Begin VB.TextBox txtLeftW 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1170
            TabIndex        =   36
            Text            =   "16"
            Top             =   3525
            Width           =   420
         End
         Begin VB.PictureBox picNormal 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   1860
            Picture         =   "frmExportHTML.frx":0024
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   35
            Top             =   2670
            Width           =   480
         End
         Begin VB.PictureBox picExpanded 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   1860
            Picture         =   "frmExportHTML.frx":00B0
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   34
            Top             =   1635
            Width           =   480
         End
         Begin VB.PictureBox picCollapsed 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   1860
            Picture         =   "frmExportHTML.frx":013D
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   33
            Top             =   585
            Width           =   480
         End
         Begin VB.TextBox txtIdent 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1170
            TabIndex        =   32
            Text            =   "20"
            Top             =   3885
            Width           =   420
         End
         Begin xfxLine3D.ucLine3D uc3DLine1 
            Height          =   30
            Left            =   135
            TabIndex        =   38
            Top             =   1230
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   53
         End
         Begin SmartButtonProject.SmartButton cmdChangeImg 
            Height          =   240
            Index           =   0
            Left            =   495
            TabIndex        =   39
            Top             =   585
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmExportHTML.frx":01CA
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
         Begin SmartButtonProject.SmartButton cmdRemoveImg 
            Height          =   240
            Index           =   0
            Left            =   495
            TabIndex        =   40
            Top             =   825
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmExportHTML.frx":0564
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
         Begin SmartButtonProject.SmartButton cmdChangeImg 
            Height          =   240
            Index           =   1
            Left            =   495
            TabIndex        =   41
            Top             =   1635
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmExportHTML.frx":08FE
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
         Begin SmartButtonProject.SmartButton cmdRemoveImg 
            Height          =   240
            Index           =   1
            Left            =   495
            TabIndex        =   42
            Top             =   1875
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmExportHTML.frx":0C98
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
         Begin SmartButtonProject.SmartButton cmdChangeImg 
            Height          =   240
            Index           =   2
            Left            =   495
            TabIndex        =   43
            Top             =   2670
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmExportHTML.frx":1032
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
         Begin SmartButtonProject.SmartButton cmdRemoveImg 
            Height          =   240
            Index           =   2
            Left            =   495
            TabIndex        =   44
            Top             =   2910
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmExportHTML.frx":13CC
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
         Begin xfxLine3D.ucLine3D uc3DLine2 
            Height          =   30
            Left            =   135
            TabIndex        =   45
            Top             =   2295
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   53
         End
         Begin xfxLine3D.ucLine3D uc3DLine3 
            Height          =   30
            Left            =   135
            TabIndex        =   46
            Top             =   3330
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   53
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "x"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1620
            TabIndex        =   52
            Top             =   3555
            Width           =   105
         End
         Begin VB.Label lblLeftS 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Icon Size"
            Height          =   195
            Left            =   255
            TabIndex        =   51
            Top             =   3555
            Width           =   645
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Normal"
            Height          =   195
            Left            =   255
            TabIndex        =   50
            Top             =   2430
            Width           =   495
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Expanded"
            Height          =   195
            Left            =   255
            TabIndex        =   49
            Top             =   1395
            Width           =   720
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Collapsed"
            Height          =   195
            Left            =   255
            TabIndex        =   48
            Top             =   345
            Width           =   690
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Indentation"
            Height          =   195
            Left            =   255
            TabIndex        =   47
            Top             =   3930
            Width           =   840
         End
      End
   End
   Begin VB.PictureBox picGeneral 
      Height          =   5850
      Left            =   195
      ScaleHeight     =   5790
      ScaleWidth      =   6525
      TabIndex        =   5
      Top             =   525
      Width           =   6585
      Begin VB.Frame Frame4 
         Caption         =   "Paths"
         Height          =   1920
         Left            =   0
         TabIndex        =   21
         Top             =   3870
         Width           =   6540
         Begin VB.OptionButton opImgPathOp 
            Height          =   225
            Index           =   1
            Left            =   510
            TabIndex        =   25
            Top             =   1485
            Width           =   255
         End
         Begin VB.OptionButton opImgPathOp 
            Caption         =   "Same as Project"
            Height          =   225
            Index           =   0
            Left            =   510
            TabIndex        =   24
            Top             =   1215
            Value           =   -1  'True
            Width           =   3720
         End
         Begin VB.TextBox txtImagesPath 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   780
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   23
            Top             =   1455
            Width           =   3990
         End
         Begin VB.TextBox txtFileName 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   510
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   555
            Width           =   4260
         End
         Begin SmartButtonProject.SmartButton cmdBrowse 
            Height          =   315
            Left            =   4845
            TabIndex        =   26
            Top             =   540
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            Picture         =   "frmExportHTML.frx":1766
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
         Begin SmartButtonProject.SmartButton cmdBrowseImagesPath 
            Height          =   315
            Left            =   4845
            TabIndex        =   27
            Top             =   1440
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            Picture         =   "frmExportHTML.frx":18C0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Images Path"
            Height          =   195
            Left            =   240
            TabIndex        =   29
            Top             =   960
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HTML File Name"
            Height          =   195
            Left            =   240
            TabIndex        =   28
            Top             =   315
            Width           =   1125
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Style"
         Height          =   2235
         Left            =   0
         TabIndex        =   11
         Top             =   1620
         Width           =   6540
         Begin VB.ComboBox cmbCommandsClass 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            TabIndex        =   17
            Text            =   "cmbCommandsClass"
            Top             =   1725
            Width           =   2130
         End
         Begin VB.ComboBox cmbGroupsClass 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1800
            TabIndex        =   16
            Text            =   "cmbGroupsClass"
            Top             =   1380
            Width           =   2130
         End
         Begin VB.TextBox txtExternalCSSFile 
            BackColor       =   &H00E0E0E0&
            Height          =   285
            Left            =   510
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            TabIndex        =   15
            Top             =   1035
            Width           =   4275
         End
         Begin VB.OptionButton opStyle 
            Caption         =   "Use External CSS File"
            Height          =   195
            Index           =   2
            Left            =   240
            TabIndex        =   14
            Top             =   825
            Width           =   3450
         End
         Begin VB.OptionButton opStyle 
            Caption         =   "Use Project Font and Color Styles"
            Height          =   195
            Index           =   1
            Left            =   240
            TabIndex        =   13
            Top             =   577
            Value           =   -1  'True
            Width           =   3420
         End
         Begin VB.OptionButton opStyle 
            Caption         =   "None"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   12
            Top             =   330
            Width           =   3420
         End
         Begin SmartButtonProject.SmartButton cmdBrowseCSSFile 
            Height          =   315
            Left            =   4845
            TabIndex        =   18
            Top             =   1020
            Width           =   360
            _ExtentX        =   635
            _ExtentY        =   556
            Picture         =   "frmExportHTML.frx":1A1A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Enabled         =   0   'False
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Commands Class"
            Height          =   195
            Left            =   510
            TabIndex        =   20
            Top             =   1785
            Width           =   1200
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Groups Class"
            Height          =   195
            Left            =   510
            TabIndex        =   19
            Top             =   1440
            Width           =   930
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Heading"
         Height          =   1590
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   6540
         Begin VB.TextBox txtTitle 
            Height          =   285
            Left            =   510
            TabIndex        =   8
            Top             =   495
            Width           =   4695
         End
         Begin VB.TextBox txtDescription 
            Height          =   285
            Left            =   510
            TabIndex        =   7
            Top             =   1080
            Width           =   4695
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Title"
            Height          =   195
            Left            =   240
            TabIndex        =   10
            Top             =   255
            Width           =   300
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description"
            Height          =   195
            Left            =   255
            TabIndex        =   9
            Top             =   840
            Width           =   795
         End
      End
   End
   Begin MSComctlLib.TabStrip tsOptions 
      Height          =   7440
      Left            =   90
      TabIndex        =   4
      Top             =   90
      Width           =   6810
      _ExtentX        =   12012
      _ExtentY        =   13123
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "General"
            Key             =   "tGeneral"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Collapsible Tree"
            Key             =   "tColTree"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview"
      Default         =   -1  'True
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   7635
      Width           =   960
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4860
      TabIndex        =   2
      Top             =   7635
      Width           =   960
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5940
      TabIndex        =   3
      Top             =   7635
      Width           =   960
   End
   Begin VB.CommandButton cmdGenerate 
      Caption         =   "Generate"
      Height          =   375
      Left            =   3615
      TabIndex        =   1
      Top             =   7635
      Width           =   960
   End
End
Attribute VB_Name = "frmExportHTML"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim OriginalHTMLSettings As ExportHTMLDef

Private Sub chkColExpAll_Click()

    UpdateControls

End Sub

Private Sub chkTree_Click()

    UpdateControls

End Sub

Private Sub cmdBrowse_Click()

    Dim ActionName As String

    With cDlg
        .DialogTitle = GetLocalizedStr(239) + " " + ActionName
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNNoReadOnlyReturn
        .filter = SupportedHTMLDocs
        .InitDir = GetRealLocal.RootWeb
        .FileName = ""
        On Error Resume Next
        .ShowOpen
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then
            Exit Sub
        End If
        txtFileName.Text = .FileName
    End With

End Sub

Private Sub cmdBrowseCSSFile_Click()

    With cDlg
        .DialogTitle = "Open HyperText Style Sheet"
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNNoReadOnlyReturn
        .filter = SupportedCSSDocs
        .InitDir = GetRealLocal.RootWeb
        .FileName = ""
        On Error Resume Next
        .ShowOpen
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then
            Exit Sub
        End If
        txtExternalCSSFile.Text = .FileName
    End With

End Sub

Private Sub cmdBrowseImagesPath_Click()

    Dim Path As String
    
    On Error Resume Next
    
    Me.Enabled = False
    
    If FolderExists(txtImagesPath.Text) Then
        Path = txtImagesPath.Text
    Else
        Path = GetRealLocal.ImagesPath
    End If
    If LenB(Dir(Path)) = 0 Or Err.number Then
        Path = ""
    Else
        Path = UnqualifyPath(Path)
    End If
    Path = BrowseForFolderByPath(Path, GetLocalizedStr(401), Me)
    
    If LenB(Path) <> 0 Then txtImagesPath.Text = AddTrailingSlash(Path, "\")
    
    Me.Enabled = True

End Sub

Private Sub cmdCancel_Click()

    Project.ExportHTMLParams = OriginalHTMLSettings
    Unload Me

End Sub

Private Sub cmdChangeImg_Click(Index As Integer)

    With cDlg
        .DialogTitle = "Open Image"
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNNoReadOnlyReturn
        .filter = SupportedImageFiles
        .InitDir = GetRealLocal.ImagesPath
        .FileName = ""
        On Error Resume Next
        .ShowOpen
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then
            Exit Sub
        End If
        
        Select Case Index
            Case 0: picCollapsed.tag = .FileName
            Case 1: picExpanded.tag = .FileName
            Case 2: picNormal.tag = .FileName
        End Select
    End With
    
    UpdateIcons

End Sub

Private Sub cmdDefaults_Click()

    picCollapsed.tag = AppPath + "exhtml\c.gif"
    picExpanded.tag = AppPath + "exhtml\o.gif"
    picNormal.tag = AppPath + "exhtml\s.gif"
    
    txtLeftW.Text = "16"
    txtLeftH.Text = "16"
    txtIdent.Text = "20"
    
    chkExpItemsLinks.Value = vbUnchecked
    chkSingleSel.Value = vbUnchecked
    
    UpdateIcons

End Sub

Private Sub cmdGenerate_Click()
    
    Dim s As exHTMLStylesConstants
    
    If opStyle(0).Value Then s = ascNone
    If opStyle(1).Value Then s = ascProject
    If opStyle(2).Value Then s = ascCSS
    
    On Error Resume Next
    SaveFile txtFileName.Text, ""
    If Err.number Then
        MsgBox "The HTML File Name '" + txtFileName.Text + "' is invalid", vbInformation + vbOKOnly, "Unable to generate Output File"
    Else
        ExportAsHTML txtTitle.Text, txtDescription.Text, txtFileName.Text, _
                s, txtExternalCSSFile.Text, cmbGroupsClass.Text, cmbCommandsClass.Text, _
                (chkTree.Value = vbChecked), _
                picCollapsed.tag, picExpanded.tag, picNormal.tag, Val(txtIdent.Text), IIf(opImgPathOp(0).Value, GetRealLocal.ImagesPath, txtImagesPath.Text), Val(txtLeftW.Text), Val(txtLeftH.Text), _
                (chkExpItemsLinks.Value = vbChecked), _
                (chkSingleSel.Value = vbChecked), _
                (chkColExpAll.Value = vbChecked), _
                txtExpAll.Text, txtColAll.Text, cmbExpColAllPlacement.ListIndex, opCodeCompliance(1).Value
    End If
    
End Sub

Private Sub UpdateIcons()

    On Error Resume Next

    picCollapsed.Picture = LoadPicture(picCollapsed.tag)
    picExpanded.Picture = LoadPicture(picExpanded.tag)
    picNormal.Picture = LoadPicture(picNormal.tag)

End Sub

Private Sub cmdOK_Click()

    With Project.ExportHTMLParams
        .Title = txtTitle.Text
        .Description = txtDescription.Text
        If opStyle(0).Value Then .Style = ascNone
        If opStyle(1).Value Then .Style = ascProject
        If opStyle(2).Value Then .Style = ascCSS
        .CSSFile = txtExternalCSSFile.Text
        .GroupClass = cmbGroupsClass.Text
        .CommandClass = cmbCommandsClass.Text
        .HTMLFileName = txtFileName.Text
        If opImgPathOp(0).Value Then .ImagesPath = "%%PROJECT%%"
        If opImgPathOp(1).Value Then .ImagesPath = txtImagesPath.Text
        .CreateTree = (chkTree.Value = vbChecked)
        .CollapsedImage = picCollapsed.tag
        .ExpandedImage = picExpanded.tag
        .NormalImage = picNormal.tag
        .IconWidth = txtLeftW.Text
        .IconHeight = txtLeftH.Text
        .Identation = txtIdent.Text
        .ExpItemsHaveLinks = (chkExpItemsLinks.Value = vbChecked)
        .SingleSelect = (chkSingleSel.Value = vbChecked)
        
        .IncludeExpCol = (chkColExpAll.Value = vbChecked)
        .ColAllStr = txtColAll.Text
        .ExpAllStr = txtExpAll.Text
        .ExpColPlacement = cmbExpColAllPlacement.ListIndex
        
        .XHTMLCompliant = opCodeCompliance(1).Value
    End With
    
    Project.HasChanged = True

    Unload Me

End Sub

Private Sub cmdPreview_Click()

    Dim s As exHTMLStylesConstants
    
    If opStyle(0).Value Then s = ascNone
    If opStyle(1).Value Then s = ascProject
    If opStyle(2).Value Then s = ascCSS
    
    ExportAsHTML txtTitle.Text, txtDescription.Text, TempPath + "smdmbp.htm", _
                s, txtExternalCSSFile.Text, cmbGroupsClass.Text, cmbCommandsClass.Text, _
                (chkTree.Value = vbChecked), _
                picCollapsed.tag, picExpanded.tag, picNormal.tag, Val(txtIdent.Text), TempPath, Val(txtLeftW.Text), Val(txtLeftH.Text), _
                (chkExpItemsLinks.Value = vbChecked), _
                (chkSingleSel.Value = vbChecked), _
                (chkColExpAll.Value = vbChecked), _
                txtExpAll.Text, txtColAll.Text, cmbExpColAllPlacement.ListIndex, opCodeCompliance(1).Value
              
    On Error Resume Next
    PreviewMode = pmcSitemap
    frmPreview.Show vbModal
    If Err.number = 400 Then
        Unload frmPreview
        DoEvents
        frmPreview.Show vbModal
    End If
    PreviewMode = pmcNormal

End Sub

Private Sub cmdRemoveImg_Click(Index As Integer)

    Select Case Index
        Case 0: picCollapsed.tag = ""
        Case 1: picExpanded.tag = ""
        Case 2: picNormal.tag = ""
    End Select
    
    UpdateIcons

End Sub

Private Sub Form_Load()

    Width = 7110
    'Height = 7485
    
    With picGeneral
        .BorderStyle = 0
        picColTree.Move .Left, .Top, .Width, .Height
        picColTree.BorderStyle = 0
        .ZOrder 0
    End With

    CenterForm Me
    
    OriginalHTMLSettings = Project.ExportHTMLParams
    
    txtTitle.Text = "<b>DHTML Menu Builder " & DMBVersion & "</b>"
    txtDescription.Text = "<small>" + Project.Name + "</small>"
    
    If FileExists(Project.FileName) And LenB(Project.ExportHTMLParams.HTMLFileName) = 0 Then
        txtFileName.Text = Replace(Project.FileName, ".dmb", ".htm")
    End If
    If txtFileName.Text = ".htm" Or Project.ExportHTMLParams.HTMLFileName = ".htm" Then
        Project.ExportHTMLParams.HTMLFileName = GetFilePath(Project.FileName) + "sitemap.htm"
    End If
    
    InitControls
    FixCtrls4Skin Me

End Sub

Private Sub InitControls()

    With Project.ExportHTMLParams
        txtTitle.Text = .Title
        txtDescription.Text = .Description
        Select Case .Style
            Case ascNone
                opStyle(0).Value = True
            Case ascProject
                opStyle(1).Value = True
            Case ascCSS
                opStyle(2).Value = True
        End Select
        txtExternalCSSFile.Text = .CSSFile
        
        GetCSSClasses
        
        cmbGroupsClass.Text = .GroupClass
        cmbCommandsClass.Text = .CommandClass
        txtFileName.Text = .HTMLFileName
        Select Case .ImagesPath
            Case "%%PROJECT%%"
                opImgPathOp(0).Value = True
            Case Else
                opImgPathOp(1).Value = True
                txtImagesPath.Text = .ImagesPath
        End Select
        chkTree.Value = IIf(.CreateTree, vbChecked, vbUnchecked)
        picCollapsed.tag = .CollapsedImage
        picExpanded.tag = .ExpandedImage
        picNormal.tag = .NormalImage
        txtLeftW.Text = .IconWidth
        txtLeftH.Text = .IconHeight
        txtIdent.Text = .Identation
        chkExpItemsLinks.Value = IIf(.ExpItemsHaveLinks, vbChecked, vbUnchecked)
        chkSingleSel.Value = IIf(.SingleSelect, vbChecked, vbUnchecked)
        
        chkColExpAll.Value = IIf(.IncludeExpCol, vbChecked, vbUnchecked)
        txtColAll.Text = .ColAllStr
        txtExpAll.Text = .ExpAllStr
        cmbExpColAllPlacement.ListIndex = .ExpColPlacement
        
        opCodeCompliance(0).Value = Not .XHTMLCompliant
        opCodeCompliance(1).Value = .XHTMLCompliant
    End With
    
    UpdateIcons
    UpdateControls

End Sub

Private Sub opImgPathOp_Click(Index As Integer)

    cmdBrowseImagesPath.Enabled = opImgPathOp(1).Value
    txtImagesPath.MousePointer = IIf(opImgPathOp(1).Value, vbDefault, vbArrow)
    
End Sub

Private Sub opStyle_Click(Index As Integer)

    UpdateControls

End Sub

Private Sub UpdateControls()

    cmdBrowseCSSFile.Enabled = txtExternalCSSFile.Enabled
    
    cmbGroupsClass.Enabled = txtExternalCSSFile.Enabled And FileExists(txtExternalCSSFile.Text)
    cmbCommandsClass.Enabled = cmbGroupsClass.Enabled
    
    txtExternalCSSFile.MousePointer = IIf(opStyle(2).Value, vbDefault, vbArrow)
    
    cmdChangeImg(0).Enabled = (chkTree.Value = vbChecked)
    cmdChangeImg(1).Enabled = (chkTree.Value = vbChecked)
    cmdChangeImg(2).Enabled = (chkTree.Value = vbChecked)
    cmdRemoveImg(0).Enabled = (chkTree.Value = vbChecked)
    cmdRemoveImg(1).Enabled = (chkTree.Value = vbChecked)
    cmdRemoveImg(2).Enabled = (chkTree.Value = vbChecked)
    
    cmdBrowseCSSFile.Enabled = (opStyle(2).Value)
    cmbGroupsClass.Enabled = (opStyle(2).Value) And FileExists(txtExternalCSSFile.Text)
    cmbCommandsClass.Enabled = (opStyle(2).Value) And FileExists(txtExternalCSSFile.Text)
    
    txtLeftW.Enabled = (chkTree.Value = vbChecked)
    txtLeftH.Enabled = (chkTree.Value = vbChecked)
    
    txtIdent.Enabled = (chkTree.Value = vbChecked)
    
    chkExpItemsLinks.Enabled = (chkTree.Value = vbChecked)
    chkSingleSel.Enabled = (chkTree.Value = vbChecked)
    
    chkColExpAll.Enabled = (chkTree.Value = vbChecked)
    txtExpAll.Enabled = (chkTree.Value = vbChecked) And (chkColExpAll.Value = vbChecked)
    txtColAll.Enabled = (chkTree.Value = vbChecked) And (chkColExpAll.Value = vbChecked)
    cmbExpColAllPlacement.Enabled = (chkTree.Value = vbChecked) And (chkColExpAll.Value = vbChecked)
    
    cmdDefaults.Enabled = (chkTree.Value = vbChecked)

End Sub

Private Sub tsOptions_Click()

    Select Case tsOptions.SelectedItem.key
        Case "tGeneral"
            picGeneral.ZOrder 0
        Case "tColTree"
            picColTree.ZOrder 0
    End Select

End Sub

Private Sub txtExternalCSSFile_Change()

    GetCSSClasses
    UpdateControls

End Sub

Private Sub GetCSSClasses()

    Dim cCode As String
    Dim l() As String
    Dim i As Integer
    Dim p As Integer
    Dim tmp As String
    
    cmbGroupsClass.Clear
    cmbCommandsClass.Clear
    
    If Not FileExists(txtExternalCSSFile.Text) Then Exit Sub
    cCode = LoadFile(txtExternalCSSFile.Text)
    
    l = Split(cCode, "{")
    
    For i = 0 To UBound(l)
        p = InStrRev(l(i), "}")
        If p = 0 Then p = 1
        tmp = Mid(l(i), p)
        tmp = Replace(tmp, " ", "")
        tmp = Replace(tmp, "}", "")
        tmp = Replace(tmp, vbTab, "")
        tmp = Replace(tmp, vbCrLf, "")
        
        If Left(tmp, 1) = "." Then
            cmbGroupsClass.AddItem tmp
            cmbCommandsClass.AddItem tmp
        End If
    Next i
    
    If cmbGroupsClass.ListCount > 0 Then
        cmbGroupsClass.ListIndex = 0
        cmbCommandsClass.ListIndex = 0
    End If

End Sub

Private Sub txtExternalCSSFile_Click()

    If opStyle(2).Value Then Exit Sub
    
    opStyle(2).Value = True
    cmdBrowseCSSFile.SetFocus

End Sub

Private Sub txtIdent_GotFocus()

    SelAll txtIdent

End Sub

Private Sub txtImagesPath_Click()

    If opImgPathOp(1).Value Then Exit Sub
    
    opImgPathOp(1).Value = True
    cmdBrowseImagesPath.SetFocus

End Sub

Private Sub txtLeftH_GotFocus()

    SelAll txtLeftH

End Sub

Private Sub txtLeftW_GotFocus()

    SelAll txtLeftW

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/exhtml.htm"

End Sub
