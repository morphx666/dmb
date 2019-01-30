VERSION 5.00
Begin VB.Form frmAddInProp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AddIn Properties"
   ClientHeight    =   2760
   ClientLeft      =   5085
   ClientTop       =   3375
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddInProp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   4335
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2325
      TabIndex        =   4
      Top             =   2310
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   5
      Top             =   2310
      Width           =   900
   End
   Begin VB.TextBox txtDesc 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   870
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1200
      Width           =   4185
   End
   Begin VB.TextBox txtName 
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
      Left            =   75
      TabIndex        =   1
      Top             =   390
      Width           =   2685
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
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
      Left            =   75
      TabIndex        =   2
      Top             =   975
      Width           =   795
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "AddIn Name"
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
      Left            =   75
      TabIndex        =   0
      Top             =   180
      Width           =   885
   End
End
Attribute VB_Name = "frmAddInProp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If LITE = 0 Then

Private Sub cmdCancel_Click()

    Project.AddIn.Name = "***"
    Unload Me

End Sub

Private Sub cmdOK_Click()

    With Project.AddIn
        If txtName.Text <> "Untitled" Then
            .Name = txtName.Text
        Else
            .Name = vbNullString
        End If
        .Description = txtDesc.Text
    End With
    
    Unload Me

End Sub

Private Sub Form_Load()

    With Project.AddIn
        If LenB(.Name) = 0 Then
            txtName.Text = "Untitled"
        Else
            txtName.Text = .Name
            txtDesc.Text = .Description
        End If
    End With

    CenterForm Me
    SetupCharset Me

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

    Dim InvChars As String
        
    InvChars = "\/:*?<>|" + Chr(34)
    If InStr(InvChars, Chr(KeyAscii)) Then
        KeyAscii = 0
        Beep
    End If

End Sub

#End If
