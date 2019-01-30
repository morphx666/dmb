VERSION 5.00
Begin VB.Form frmParamSel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Select Parameter"
   ClientHeight    =   3165
   ClientLeft      =   6885
   ClientTop       =   5850
   ClientWidth     =   3495
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3495
   ShowInTaskbar   =   0   'False
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
      Left            =   2505
      TabIndex        =   2
      Top             =   2715
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Enabled         =   0   'False
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
      Left            =   2505
      TabIndex        =   1
      Top             =   2265
      Width           =   900
   End
   Begin VB.ListBox lstParams 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3000
      IntegralHeight  =   0   'False
      Left            =   60
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   90
      Width           =   2310
   End
End
Attribute VB_Name = "frmParamSel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    With frmAddInEditor.txtCodeEditor(0)
        .SelText = vbNullString
        .Text = Left(.Text, .SelStart) + "??" + lstParams.List(lstParams.ListIndex) + "??" + Mid(.Text, .SelStart + 1)
    End With
    
    Unload Me

End Sub

Private Sub Form_Load()

    Dim i As Integer

    CenterForm Me
    SetupCharset Me
    
    For i = 1 To UBound(params)
        lstParams.AddItem params(i).Name
    Next i

End Sub

Private Sub lstParams_Click()

    cmdOK.Enabled = True

End Sub

Private Sub lstParams_DblClick()

    If lstParams.ListIndex <> -1 Then cmdOK_Click

End Sub
