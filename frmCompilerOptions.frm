VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmCompilerOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compiler Options"
   ClientHeight    =   3840
   ClientLeft      =   6690
   ClientTop       =   5430
   ClientWidth     =   4815
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompilerOptions.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3840
   ScaleWidth      =   4815
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Caption         =   "Special Options"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   675
      TabIndex        =   8
      Top             =   1455
      Width           =   3750
      Begin xfxLine3D.ucLine3D uc3DLine2 
         Height          =   30
         Left            =   165
         TabIndex        =   13
         Top             =   1140
         Width           =   3345
         _ExtentX        =   5900
         _ExtentY        =   53
      End
      Begin VB.CheckBox chkSharedProject 
         Caption         =   "This is a Shared Project"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   12
         Top             =   1260
         Width           =   2970
      End
      Begin VB.CheckBox chkhRefFile 
         Caption         =   "Generate hRef file"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   11
         Top             =   825
         Width           =   2970
      End
      Begin VB.CheckBox chkNSCode 
         Caption         =   "Generate Code for Navigator 4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   10
         Top             =   570
         Width           =   2970
      End
      Begin VB.CheckBox chkIECode 
         Caption         =   "Generate Code for DOM compliant browsers"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   165
         TabIndex        =   9
         Top             =   315
         Width           =   3555
      End
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   45
      TabIndex        =   7
      Top             =   3240
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   53
   End
   Begin VB.CheckBox chkShowReport 
      Caption         =   "Show Report"
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
      Left            =   120
      TabIndex        =   2
      Top             =   3495
      Width           =   1665
   End
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Default         =   -1  'True
      Height          =   375
      Left            =   2775
      TabIndex        =   3
      Top             =   3405
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3855
      TabIndex        =   4
      Top             =   3405
      Width           =   900
   End
   Begin VB.ComboBox cmbConfiguration 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCompilerOptions.frx":058A
      Left            =   810
      List            =   "frmCompilerOptions.frx":058C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   990
      WhatsThisHelpID =   20400
      Width           =   2850
   End
   Begin VB.ComboBox cmbCodeOp 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmCompilerOptions.frx":058E
      Left            =   810
      List            =   "frmCompilerOptions.frx":0598
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   285
      WhatsThisHelpID =   20400
      Width           =   2850
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   135
      Picture         =   "frmCompilerOptions.frx":05D2
      Top             =   585
      Width           =   240
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuration"
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
      Left            =   810
      TabIndex        =   6
      Top             =   780
      Width           =   975
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Code Optimization"
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
      Left            =   810
      TabIndex        =   5
      Top             =   75
      Width           =   1305
   End
End
Attribute VB_Name = "frmCompilerOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    AbortCompileDlg = True
    Unload Me

End Sub

Private Sub cmdCompile_Click()

    ShowReport = (chkShowReport.Value = vbChecked)

    Project.CodeOptimization = cmbCodeOp.ListIndex
    Project.DefaultConfig = cmbConfiguration.ListIndex
    
    Project.CompileIECode = (chkIECode.Value = vbChecked)
    Project.CompileNSCode = (chkNSCode.Value = vbChecked)
    Project.CompilehRefFile = (chkhRefFile.Value = vbChecked)
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 112 Then ShowHelp "dialogs/compiler_options.htm"

End Sub

Private Sub Form_Load()

    Dim i As Integer
    Dim UsesFrames As Boolean

    CenterForm Me
    
    AbortCompileDlg = False
    
    cmbCodeOp.ListIndex = Project.CodeOptimization
    
    For i = 0 To UBound(Project.UserConfigs)
        cmbConfiguration.AddItem Project.UserConfigs(i).Name
    Next i
    cmbConfiguration.ListIndex = Project.DefaultConfig
    
    chkIECode.Value = Abs(Project.CompileIECode)
    chkNSCode.Value = Abs(Project.CompileNSCode)
    chkhRefFile.Value = Abs(Project.CompilehRefFile)
    
    UsesFrames = Project.UserConfigs(Project.DefaultConfig).Frames.UseFrames

End Sub
