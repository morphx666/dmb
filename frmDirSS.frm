VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmDirSS 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Directory Structure Scanner"
   ClientHeight    =   5850
   ClientLeft      =   9255
   ClientTop       =   4170
   ClientWidth     =   5745
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDirSS.frx":0000
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDummy 
      Height          =   285
      Left            =   195
      TabIndex        =   18
      TabStop         =   0   'False
      Text            =   "Text1"
      Top             =   5280
      Visible         =   0   'False
      Width           =   3285
   End
   Begin VB.PictureBox picFlood 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   195
      ScaleHeight     =   255
      ScaleWidth      =   5325
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4815
      Width           =   5355
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Start"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3585
      TabIndex        =   19
      Top             =   5325
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4650
      TabIndex        =   20
      Top             =   5325
      Width           =   900
   End
   Begin VB.Frame Frame3 
      Caption         =   "Build Menus Structure From"
      Height          =   1485
      Left            =   195
      TabIndex        =   3
      Top             =   885
      Width           =   2910
      Begin VB.CheckBox chkGroupRootDocs 
         Caption         =   "Group Documents in Root"
         Height          =   195
         Left            =   495
         TabIndex        =   5
         Top             =   570
         Value           =   1  'Checked
         Width           =   2250
      End
      Begin VB.OptionButton opScan 
         Caption         =   "Directory Structure"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   4
         Top             =   330
         Value           =   -1  'True
         Width           =   2505
      End
      Begin VB.OptionButton opScan 
         Caption         =   "Links in Documents"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   930
         Width           =   2310
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "File Types"
      Height          =   1980
      Left            =   195
      TabIndex        =   11
      Top             =   2475
      Width           =   5355
      Begin VB.CommandButton cmdDefault 
         Caption         =   "Default"
         Height          =   360
         Left            =   4245
         TabIndex        =   15
         Top             =   1425
         Width           =   915
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add..."
         Height          =   360
         Left            =   4245
         TabIndex        =   13
         Top             =   330
         Width           =   915
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   360
         Left            =   4245
         TabIndex        =   14
         Top             =   802
         Width           =   915
      End
      Begin MSComctlLib.ListView lvFileTypes 
         Height          =   1455
         Left            =   240
         TabIndex        =   12
         Top             =   330
         Width           =   3870
         _ExtentX        =   6826
         _ExtentY        =   2566
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Key             =   "chExt"
            Text            =   "Extension"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Key             =   "chDesc"
            Text            =   "Description"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Create Menu Items From"
      Height          =   1485
      Left            =   3225
      TabIndex        =   7
      Top             =   885
      Width           =   2325
      Begin VB.OptionButton opCreate 
         Caption         =   "Documents' Title"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   330
         Value           =   -1  'True
         Width           =   1920
      End
      Begin VB.OptionButton opCreate 
         Caption         =   "File Name"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   630
         Width           =   1920
      End
      Begin VB.OptionButton opCreate 
         Caption         =   "Link Text"
         Enabled         =   0   'False
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   930
         Width           =   1920
      End
   End
   Begin VB.TextBox txtDir 
      Height          =   285
      Left            =   195
      TabIndex        =   1
      Top             =   450
      Width           =   3570
   End
   Begin SmartButtonProject.SmartButton cmdBrowse 
      Height          =   315
      Left            =   3840
      TabIndex        =   2
      Top             =   435
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmDirSS.frx":058A
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
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   4545
      Top             =   225
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "0"
      DialogTitle     =   "Select Project"
      Filter          =   "DHTML Menu Builder Projects|*.dmb"
      MaxFileSize     =   1024
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Progress"
      Height          =   195
      Left            =   195
      TabIndex        =   16
      Top             =   4575
      Width           =   630
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Root Document"
      Height          =   195
      Left            =   195
      TabIndex        =   0
      Top             =   210
      Width           =   1110
   End
End
Attribute VB_Name = "frmDirSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    SetValidFileTypes
    frmDirSS_AddFileType.Show vbModal
    LoadFileTypes

End Sub

Private Sub cmdBrowse_Click()

    SetValidFileTypes

    With cDlg
        .DialogTitle = "Select Root Document"
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNNoReadOnlyReturn
        .filter = "Selected File Types|" + DirSSValidFileTypes + "|All Files|*.*"
        .FileName = ""
        On Error Resume Next
        .ShowOpen
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then
            Exit Sub
        End If
        txtDir.Text = .FileName
    End With

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    DelFileTypes

End Sub

Private Sub cmdDefault_Click()

    SaveSetting "DMB", "Preferences", "CustomFileTypes", strGetSupportedHTMLDocs
    LoadFileTypes

End Sub

Private Sub cmdOK_Click()

    Me.Enabled = False

    SetValidFileTypes
    StartScan
    
    Unload Me

End Sub

Private Sub SetValidFileTypes()

    'DirSSValidFileTypes = "*." + Join(Split(strGetSupportedHTMLDocs, ";"), ";*.")
    Dim nItem As ListItem
    
    DirSSValidFileTypes = ""
    For Each nItem In lvFileTypes.ListItems
        DirSSValidFileTypes = DirSSValidFileTypes + "*." + nItem.Text + ";"
    Next nItem
    
    SaveSetting "DMB", "Preferences", "CustomFileTypes", Replace(Join(Split(DirSSValidFileTypes, ";")), "*.", ";")

End Sub

Private Sub StartScan()

    Dim sTextMode As dirssTextMode
    
    LockWindowUpdate frmMain.tvMenus.hwnd
        
    If opCreate(0).Value Then sTextMode = tmDocTitle
    If opCreate(1).Value Then sTextMode = tmFileName
    If opCreate(2).Value Then sTextMode = tmLinkText

    If opScan(0).Value Then
        StartDirScan GetFilePath(txtDir.Text), sTextMode, (chkGroupRootDocs.Value = vbChecked)
    End If
    
    If opScan(1).Value Then
        StartLinksScan txtDir.Text, sTextMode
    End If
    
    LockWindowUpdate 0

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/dirss.htm"

End Sub

Private Sub Form_Load()

    'Width = 5865
    'Height = 6285

    CenterForm Me
    FixCtrls4Skin Me
    
    LoadFileTypes

End Sub

Private Sub DelFileTypes()

    Dim sItem As ListItem
    Dim s As String
    
ReStart:
    With lvFileTypes
        For Each sItem In .ListItems
            If sItem.Selected Then
                .ListItems.Remove sItem.Index
                GoTo ReStart
            End If
        Next sItem
        
        For Each sItem In .ListItems
            s = s + sItem.Text + ";"
        Next sItem
        SaveSetting "DMB", "Preferences", "CustomFileTypes", s
    End With
    
End Sub

Private Sub LoadFileTypes()

    Dim FT() As String
    Dim i As Integer
    
    On Error Resume Next
    
    FT = Split(GetSetting("DMB", "Preferences", "CustomFileTypes", ""), ";")
    If UBound(FT) = -1 Then
        FT = Split(strGetSupportedHTMLDocs, ";")
    End If

    lvFileTypes.MultiSelect = False
    With lvFileTypes.ListItems
        .Clear
        
        For i = 0 To UBound(FT)
            FT(i) = Trim(FT(i))
            If LenB(FT(i)) <> 0 Then
                .Add(, "K" + FT(i), FT(i)).SubItems(1) = GetExtDesc(FT(i))
            End If
        Next i
        
        If .Count > 0 Then
            .item(1).Selected = True
            .item(1).EnsureVisible
        End If
    End With
    lvFileTypes.MultiSelect = True
    
    CoolListView lvFileTypes

End Sub

Private Function GetExtDesc(ext As String) As String

    Dim d As String
    Dim dd As String
    
    d = QueryValue(HKEY_CLASSES_ROOT, "." + ext)
    dd = QueryValue(HKEY_CLASSES_ROOT, d)
    
    GetExtDesc = IIf(LenB(dd) = 0, d, dd)

End Function

Private Sub opScan_Click(Index As Integer)

    opCreate(2).Enabled = opScan(1).Value
    chkGroupRootDocs.Enabled = opScan(0).Value

End Sub

Private Sub txtDir_Change()

    cmdOK.Enabled = FileExists(txtDir.Text)

End Sub
