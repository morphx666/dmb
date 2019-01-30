VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmConfigSelect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Configuration"
   ClientHeight    =   2985
   ClientLeft      =   7455
   ClientTop       =   5925
   ClientWidth     =   4080
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
   ScaleHeight     =   2985
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3075
      TabIndex        =   4
      Top             =   2520
      Width           =   900
   End
   Begin VB.ListBox lstConfigs 
      Height          =   1860
      Left            =   105
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   345
      Width           =   3870
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   2
      Top             =   2385
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   53
   End
   Begin VB.Label lblConfigurations 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configurations"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   105
      Width           =   1050
   End
End
Attribute VB_Name = "frmConfigSelect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsUpdating As Boolean

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    If Project.DefaultConfig <> lstConfigs.ListIndex Then
        DisplayTip GetLocalizedStr(682), GetLocalizedStr(683)
    End If

    Project.DefaultConfig = lstConfigs.ListIndex
    Unload Me

End Sub

Private Sub Form_Load()

    Dim i As Integer
    
    LocalizeUI
    SetupCharset Me
    CenterForm Me
    
    For i = 0 To UBound(Project.UserConfigs)
        lstConfigs.AddItem Project.UserConfigs(i).Name + " (" + ConfigTypeName(Project.UserConfigs(i)) + ")"
        lstConfigs.Selected(lstConfigs.NewIndex) = (Project.DefaultConfig = i)
    Next i

End Sub

Private Sub lstConfigs_Click()

    If IsUpdating Then Exit Sub
    lstConfigs_ItemCheck lstConfigs.ListIndex

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

Private Sub LocalizeUI()

    Caption = GetLocalizedStr(681)

    lblConfigurations.Caption = GetLocalizedStr(322)
    
    cmdOK.Caption = GetLocalizedStr(186)
    cmdCancel.Caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
