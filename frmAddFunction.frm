VERSION 5.00
Begin VB.Form frmAddFunction 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Function"
   ClientHeight    =   2760
   ClientLeft      =   7965
   ClientTop       =   6555
   ClientWidth     =   3405
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddFunction.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   3405
   ShowInTaskbar   =   0   'False
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
      Left            =   60
      TabIndex        =   1
      Top             =   285
      Width           =   3270
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
      Left            =   60
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   3270
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
      Left            =   2430
      TabIndex        =   5
      Top             =   2295
      Width           =   900
   End
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
      Left            =   1410
      TabIndex        =   4
      Top             =   2295
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Function Name"
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
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   1065
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
      Left            =   60
      TabIndex        =   2
      Top             =   840
      Width           =   795
   End
End
Attribute VB_Name = "frmAddFunction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim i As Integer

    If cmdOK.Caption = "&Update" Then
        For i = 1 To UBound(Sections)
            If Sections(i).Name = txtName.Tag Then
                Sections(i).Name = txtName.Text
                Sections(i).Description = txtDesc.Text
                'Sections(i).Support = IIf(chkIE.Value = vbChecked, "IE", "")
                'Sections(i).Support = Sections(i).Support + IIf(chkNS.Value = vbChecked, ",NS", "")
                Exit For
            End If
        Next i
    Else
        ReDim Preserve Sections(UBound(Sections) + 1)
        With Sections(UBound(Sections))
            .Name = txtName.Text
            .code = vbNullString
            .NewCode = vbNullString
            .params = vbNullString
            .NewParams = vbNullString
            .Description = txtDesc.Text + vbCrLf
            '.Support = IIf(chkIE.Value = vbChecked, "IE", "")
            '.Support = .Support + IIf(chkNS.Value = vbChecked, ",NS", "")
            .IsNew = True
        End With
    End If
    
    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me

End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)

    If Not ((KeyAscii >= Asc("a") And KeyAscii <= Asc("z")) Or _
            (KeyAscii >= Asc("A") And KeyAscii <= Asc("Z")) Or _
            (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Or _
            KeyAscii = Asc(" ") Or KeyAscii = Asc("_") Or KeyAscii = 8) Then
        KeyAscii = 0
        Beep
    End If

End Sub
