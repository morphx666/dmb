VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmPresetCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Preset Properties"
   ClientHeight    =   4020
   ClientLeft      =   6375
   ClientTop       =   5805
   ClientWidth     =   4170
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCreatePreset.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4020
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.FileListBox flbPresets 
      Height          =   1065
      Left            =   2595
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   765
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.TextBox txtTitle 
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   390
      Width           =   3225
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   60
      TabIndex        =   9
      Top             =   3390
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   2175
      TabIndex        =   10
      Top             =   3570
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3195
      TabIndex        =   11
      Top             =   3570
      Width           =   900
   End
   Begin VB.ComboBox cmbCategory 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Text            =   "Combo1"
      Top             =   2895
      Width           =   3225
   End
   Begin VB.TextBox txtComments 
      Height          =   795
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   1725
      Width           =   3225
   End
   Begin VB.TextBox txtAuthor 
      Enabled         =   0   'False
      Height          =   285
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1050
      Width           =   3225
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   165
      Width           =   300
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Category"
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2655
      Width           =   675
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Comments"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   1485
      Width           =   750
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Author"
      Height          =   195
      Left            =   120
      TabIndex        =   2
      Top             =   825
      Width           =   495
   End
End
Attribute VB_Name = "frmPresetCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbCategory_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc("|") Then
        KeyAscii = 0
    End If

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdSave_Click()

    Dim sFile As String
    
    sFile = AppPath + "Presets\" + txtTitle.Text + ".dpp"
    
    If Not IsInIDE Then txtAuthor.Text = USER

    If FileExists(sFile) Then
        If MsgBox(GetLocalizedStr(895) + " '" + txtTitle.Text + "' " + GetLocalizedStr(896), vbQuestion + vbYesNo, GetLocalizedStr(834)) = vbNo Then GoTo ExitSub
        If GetPresetProperty(sFile, piAuthor) <> txtAuthor.Text Then
            MsgBox GetLocalizedStr(897), vbInformation + vbOKOnly, GetLocalizedStr(898)
            GoTo ExitSub
        End If
    End If
    
    CompressPreset txtTitle.Text, txtAuthor.Text, txtComments.Text, cmbCategory.Text, flbPresets
    
ExitSub:
    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    
    cmbCategory.AddItem "(uncategorized)"
    cmbCategory.AddItem "Simple"
    cmbCategory.AddItem "Horizontal"
    cmbCategory.AddItem "Effects"
    cmbCategory.AddItem "Multiple Toolbars"
    cmbCategory.AddItem "Images"
    cmbCategory.AddItem "Professional"
    cmbCategory.AddItem "Operating Systems"
    cmbCategory.AddItem "Applications"
    
    cmbCategory.ListIndex = 0
    
    txtTitle.Text = Project.Name
    txtAuthor.Text = USER
    
    If IsInIDE Then
        txtAuthor.Enabled = True
        txtAuthor.Locked = False
    End If

End Sub

Private Sub txtAuthor_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc("|") Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtComments_KeyPress(KeyAscii As Integer)

    If KeyAscii = Asc("|") Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtTitle_KeyPress(KeyAscii As Integer)

    If (KeyAscii < Asc("a") Or KeyAscii > Asc("z")) And _
        (KeyAscii < Asc("A") Or KeyAscii > Asc("Z")) And _
        (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And _
        KeyAscii <> 8 And KeyAscii <> Asc("-") And KeyAscii <> Asc("_") And KeyAscii <> 32 Then
        KeyAscii = 0
    End If

End Sub
