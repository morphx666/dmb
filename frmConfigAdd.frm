VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmConfigAdd 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Configuration"
   ClientHeight    =   3735
   ClientLeft      =   7005
   ClientTop       =   5205
   ClientWidth     =   4335
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
   ScaleHeight     =   3735
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   8
      Top             =   3120
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   53
   End
   Begin VB.ComboBox cmbConfigs 
      Height          =   315
      ItemData        =   "frmConfigAdd.frx":0000
      Left            =   90
      List            =   "frmConfigAdd.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2475
      Width           =   1830
   End
   Begin VB.ComboBox cmbType 
      Height          =   315
      ItemData        =   "frmConfigAdd.frx":0012
      Left            =   90
      List            =   "frmConfigAdd.frx":001F
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   1035
      Width           =   1830
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   375
      Width           =   2460
   End
   Begin VB.TextBox txtDesc 
      Height          =   315
      Left            =   90
      TabIndex        =   5
      Top             =   1740
      Width           =   4020
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3375
      TabIndex        =   10
      Top             =   3285
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   2325
      TabIndex        =   9
      Top             =   3285
      Width           =   900
   End
   Begin VB.Label lblCreateFrom 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Create From"
      Height          =   195
      Left            =   90
      TabIndex        =   6
      Top             =   2250
      Width           =   900
   End
   Begin VB.Label lblType 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Type"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   810
      Width           =   360
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   405
   End
   Begin VB.Label lblDescription 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      Height          =   195
      Left            =   90
      TabIndex        =   4
      Top             =   1500
      Width           =   795
   End
End
Attribute VB_Name = "frmConfigAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim SelConfig As Integer
    Dim i As Integer
    
    If Caption = GetLocalizedStr(394) Then
        SelConfig = CInt(Tag)
    Else
        For i = 0 To UBound(Project.UserConfigs)
            If Project.UserConfigs(i).Name = txtName.Text Then
                MsgBox "The project already contains a configuration named " + txtName.Text + vbCrLf + "Please choose a different name for this new configuration", vbInformation + vbOKOnly, "Error Creating New Configuration"
                Exit Sub
            End If
        Next i
    
        ReDim Preserve Project.UserConfigs(UBound(Project.UserConfigs) + 1)
        If cmbConfigs.ListIndex > 0 Then
            Project.UserConfigs(UBound(Project.UserConfigs)) = Project.UserConfigs(cmbConfigs.ListIndex - 1)
        End If
        SelConfig = UBound(Project.UserConfigs)
    End If
    
    With Project.UserConfigs(SelConfig)
        .Name = txtName.Text
        .Description = txtDesc.Text
        Select Case cmbType.ListIndex
            Case 0
                .Type = ctcLocal
            Case 1
                .Type = ctcRemote
                .OptmizePaths = True
            Case 2
                .Type = ctcCDROM
        End Select
        If cmbConfigs.ListIndex > 0 Then
            If .Type <> Project.UserConfigs(cmbConfigs.ListIndex - 1).Type Then
                .RootWeb = vbNullString
                .CompiledPath = vbNullString
                .ImagesPath = vbNullString
                If .Type = ctcRemote And Project.UserConfigs(cmbConfigs.ListIndex - 1).Type = ctcLocal Then
                    .LocalInfo4RemoteConfig = Project.UserConfigs(cmbConfigs.ListIndex - 1).Name
                End If
            End If
        End If
    End With
    
    Unload Me

End Sub

Private Sub Form_Load()

    Dim i As Integer

    SetupCharset Me
    LocalizeUI
    CenterForm Me
    
    cmbType.ListIndex = 2
    
    For i = 0 To UBound(Project.UserConfigs)
        cmbConfigs.AddItem Project.UserConfigs(i).Name
    Next i
    cmbConfigs.ListIndex = 0

End Sub

Private Sub txtName_Change()

    cmdOK.Enabled = (LenB(txtName.Text) <> 0)

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

    'If Not ((KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or _
    '        (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or _
    '        (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or _
    '        KeyAscii = Asc(" ") Or KeyAscii = Asc("_") Or _
    '        KeyAscii = "(" Or KeyAscii = ")" Or KeyAscii = 8) Then
    '    KeyAscii = 0
    '    Beep
    'End If

End Sub

Private Sub LocalizeUI()

    lblName.Caption = GetLocalizedStr(409)
    lblType.Caption = GetLocalizedStr(711)
    lblCreateFrom.Caption = GetLocalizedStr(712)

    cmdOK.Caption = GetLocalizedStr(186)
    cmdCancel.Caption = GetLocalizedStr(187)

End Sub
