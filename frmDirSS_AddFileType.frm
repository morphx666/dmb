VERSION 5.00
Begin VB.Form frmDirSS_AddFileType 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add File Type"
   ClientHeight    =   1710
   ClientLeft      =   6660
   ClientTop       =   3765
   ClientWidth     =   4440
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1710
   ScaleWidth      =   4440
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3435
      TabIndex        =   4
      Top             =   1230
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2385
      TabIndex        =   3
      Top             =   1230
      Width           =   900
   End
   Begin VB.TextBox txtExt 
      Height          =   285
      Left            =   180
      TabIndex        =   1
      Top             =   450
      Width           =   2775
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Separate multiple extensions with ;"
      ForeColor       =   &H00808080&
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   750
      Width           =   2520
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "New File Extension"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   210
      Width           =   1350
   End
End
Attribute VB_Name = "frmDirSS_AddFileType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    CenterForm Me

    Unload Me

End Sub

Private Sub cmdOK_Click()

    If LenB(txtExt.Text) <> 0 Then
        txtExt.Text = Replace(txtExt.Text, "*", "")
        txtExt.Text = Replace(txtExt.Text, ".", "")
        SaveSetting "DMB", "Preferences", "CustomFileTypes", Replace(GetSetting("DMB", "Preferences", "CustomFileTypes", "") + ";" + UCase(txtExt.Text), ";;", ";")
    End If
    
    Unload Me

End Sub
