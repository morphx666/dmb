VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmErrValidation 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Validation Unsuccessful"
   ClientHeight    =   4815
   ClientLeft      =   5775
   ClientTop       =   3720
   ClientWidth     =   6165
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
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
   ScaleHeight     =   4815
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtServerResponse 
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   45
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      TabStop         =   0   'False
      Text            =   "frmErrValidation.frx":0000
      Top             =   1230
      Width           =   6075
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   6
      Top             =   4290
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   345
      Left            =   4920
      TabIndex        =   7
      Top             =   4410
      Width           =   1200
   End
   Begin VB.TextBox txtMsg 
      BackColor       =   &H00C0FFFF&
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
      Left            =   45
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "frmErrValidation.frx":0023
      Top             =   2595
      Width           =   6075
   End
   Begin xfxLine3D.ucLine3D uc3DLine2 
      Height          =   30
      Left            =   0
      TabIndex        =   1
      Top             =   930
      Width           =   6120
      _ExtentX        =   10795
      _ExtentY        =   53
   End
   Begin VB.Label lblAdInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Additional Information:"
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
      Left            =   45
      TabIndex        =   4
      Top             =   2370
      Width           =   1650
   End
   Begin VB.Label lblServerResponse 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The reason reported by the server is:"
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
      Left            =   45
      TabIndex        =   2
      Top             =   1005
      Width           =   2715
   End
   Begin VB.Label lblTopMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "The Server has not accepted the provided registration information."
      Height          =   420
      Left            =   720
      TabIndex        =   0
      Top             =   255
      Width           =   5280
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   105
      Picture         =   "frmErrValidation.frx":01BD
      Top             =   225
      Width           =   480
   End
End
Attribute VB_Name = "frmErrValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    
    LocalizeUI
    
    frmErrDlgIsVisible = True

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    frmErrDlgIsVisible = False

End Sub

Private Sub txtMsg_GotFocus()

    cmdCancel.SetFocus

End Sub

Private Sub LocalizeUI()

    Caption = GetLocalizedStr(645)
    
    lblTopMsg.Caption = GetLocalizedStr(648)
    lblServerResponse.Caption = GetLocalizedStr(649)
    lblAdInfo.Caption = GetLocalizedStr(650)
    
    cmdCancel.Caption = GetLocalizedStr(424)

End Sub
