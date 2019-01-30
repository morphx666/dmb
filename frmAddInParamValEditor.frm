VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmAddInParamValEditor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "AddIn Parameters"
   ClientHeight    =   3480
   ClientLeft      =   4080
   ClientTop       =   6270
   ClientWidth     =   4260
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
   ScaleHeight     =   3480
   ScaleWidth      =   4260
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDefault 
      Caption         =   "Default"
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
      Left            =   3165
      TabIndex        =   5
      Top             =   1995
      Width           =   705
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
      Left            =   2265
      TabIndex        =   7
      Top             =   3045
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
      Left            =   3300
      TabIndex        =   8
      Top             =   3045
      Width           =   900
   End
   Begin VB.TextBox txtValue 
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
      TabIndex        =   4
      Top             =   1995
      Width           =   3000
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   15
      TabIndex        =   2
      Top             =   765
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   53
   End
   Begin VB.ComboBox cmbParams 
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
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   315
      Width           =   3000
   End
   Begin xfxLine3D.ucLine3D uc3DLine2 
      Height          =   30
      Left            =   15
      TabIndex        =   6
      Top             =   2880
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   53
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Description here"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   660
      Left            =   60
      TabIndex        =   3
      Top             =   1035
      UseMnemonic     =   0   'False
      Width           =   4065
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parameter Name"
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
      Top             =   90
      Width           =   1200
   End
End
Attribute VB_Name = "frmAddInParamValEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If LITE = 0 Then

Dim NewParams() As AddInParameter
Dim SelParamIdx As Integer
Dim IsUpdating As Boolean

Private Sub cmbParams_Click()

    Dim i As Integer
    
    IsUpdating = True

    For i = 1 To UBound(NewParams)
        If NewParams(i).Name = cmbParams.Text Then
            SelParamIdx = i
            Exit For
        End If
    Next i
    
    txtValue.Text = NewParams(SelParamIdx).Value
    lblDescription.Caption = NewParams(SelParamIdx).Description
    cmdDefault.Enabled = (LenB(NewParams(SelParamIdx).Default) <> 0)
    
    IsUpdating = False

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdDefault_Click()

    txtValue.Text = NewParams(SelParamIdx).Default

End Sub

Private Sub cmdOK_Click()

    params = NewParams
    
    Unload Me

End Sub

Private Sub Form_Load()

    Dim i As Integer

    CenterForm Me
    SetupCharset Me
    
    NewParams = params
    
    For i = 1 To UBound(NewParams)
        cmbParams.AddItem NewParams(i).Name
    Next i
    
    cmbParams.ListIndex = 0

End Sub

Private Sub txtValue_Change()

    If IsUpdating Then Exit Sub
    
    NewParams(SelParamIdx).Value = txtValue.Text

End Sub

#End If
