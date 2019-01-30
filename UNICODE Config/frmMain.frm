VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DHTML Menu Builder UNICODE Config"
   ClientHeight    =   1350
   ClientLeft      =   7185
   ClientTop       =   8580
   ClientWidth     =   4665
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1350
   ScaleWidth      =   4665
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   375
      Left            =   3555
      TabIndex        =   0
      Top             =   915
      Width           =   1050
   End
   Begin VB.CheckBox chkToggle 
      Caption         =   "Enable DHTML Menu Builder UNICODE Support"
      Height          =   420
      Left            =   195
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   4155
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkToggle_Click()

    SaveSetting "DMB", "Preferences", "DoUNICODE", IIf(chkToggle.Value = vbChecked, 1, 0)

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    With Screen
        Move (.Width - Width) \ 2, (.Height - Height) \ 2
    End With
    
    chkToggle.Value = IIf(GetSetting("DMB", "Preferences", "DoUNICODE", 0) = 0, 0, 1)

End Sub
