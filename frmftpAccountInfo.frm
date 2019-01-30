VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmFTPAccountInfo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Account Information"
   ClientHeight    =   2130
   ClientLeft      =   7200
   ClientTop       =   6840
   ClientWidth     =   3810
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
   Icon            =   "frmftpAccountInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   3810
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   540
      Top             =   1560
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2865
      TabIndex        =   6
      Top             =   1695
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1815
      TabIndex        =   5
      Top             =   1695
      Width           =   900
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   15
      TabIndex        =   4
      Top             =   1545
      Width           =   3840
      _ExtentX        =   6773
      _ExtentY        =   53
   End
   Begin VB.TextBox txtPassword 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   765
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1050
      Width           =   2820
   End
   Begin VB.TextBox txtUserName 
      Height          =   285
      Left            =   765
      TabIndex        =   1
      Top             =   360
      Width           =   2820
   End
   Begin VB.Label lblPassword 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      Height          =   195
      Left            =   765
      TabIndex        =   2
      Top             =   810
      Width           =   690
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   135
      Picture         =   "frmftpAccountInfo.frx":27A2
      Stretch         =   -1  'True
      Top             =   570
      Width           =   480
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   765
      TabIndex        =   0
      Top             =   120
      Width           =   780
   End
End
Attribute VB_Name = "frmFTPAccountInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    ftpUserName = ""
    ftpPassword = ""
    
    Unload Me

End Sub

Private Sub cmdOK_Click()

    ftpUserName = txtUserName
    ftpPassword = txtPassword
    
    Unload Me

End Sub

Private Sub Form_Load()

    LocalizeUI
    
    On Error Resume Next
    
    CenterForm Me
    tmrInit.Enabled = True

End Sub

Private Sub LocalizeUI()

    Caption = GetLocalizedStr(677)
    
    lblUserName.Caption = GetLocalizedStr(373)
    lblPassword.Caption = GetLocalizedStr(374)
    
    cmdOK.Caption = GetLocalizedStr(186)
    cmdCancel.Caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub

Private Sub tmrInit_Timer()

    tmrInit.Enabled = False
    
    If txtUserName.Text <> "" Then txtPassword.SetFocus

End Sub
