VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmAIEConfirm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AddIn Editor"
   ClientHeight    =   2040
   ClientLeft      =   6060
   ClientTop       =   6690
   ClientWidth     =   5670
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
   ScaleHeight     =   2040
   ScaleWidth      =   5670
   ShowInTaskbar   =   0   'False
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   45
      TabIndex        =   1
      Top             =   1440
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   53
   End
   Begin VB.CheckBox chkDontAsk 
      Caption         =   "Don't ask me again"
      Height          =   195
      Left            =   45
      TabIndex        =   2
      Top             =   1515
      Width           =   1965
   End
   Begin VB.CommandButton cmdYes 
      Caption         =   "Yes"
      Default         =   -1  'True
      Height          =   375
      Left            =   3660
      TabIndex        =   3
      Top             =   1590
      Width           =   900
   End
   Begin VB.CommandButton cmdNo 
      Cancel          =   -1  'True
      Caption         =   "No"
      Height          =   375
      Left            =   4695
      TabIndex        =   4
      Top             =   1590
      Width           =   900
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   225
      Picture         =   "frmAIEConfirm.frx":0000
      Top             =   360
      Width           =   480
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAIEConfirm.frx":0742
      Height          =   780
      Left            =   960
      TabIndex        =   0
      Top             =   210
      Width           =   4485
   End
End
Attribute VB_Name = "frmAIEConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If LITE = 0 Then

Private Sub chkDontAsk_Click()

    Preferences.ShowWarningAddInEditor = Not (chkDontAsk.Value = vbChecked)

End Sub

Private Sub cmdNo_Click()

    DlgAns = vbNo
    Unload Me

End Sub

Private Sub cmdYes_Click()

    DlgAns = vbYes
    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    
    LocalizeUI

End Sub

Private Sub LocalizeUI()

    lblInfo.Caption = GetLocalizedStr(415)
    
    chkDontAsk.Caption = GetLocalizedStr(418)

    cmdYes.Caption = GetLocalizedStr(416)
    cmdNo.Caption = GetLocalizedStr(417)
    
    FixContolsWidth Me

End Sub

#End If
