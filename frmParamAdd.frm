VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmParamMan 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Parameter"
   ClientHeight    =   3480
   ClientLeft      =   6465
   ClientTop       =   5880
   ClientWidth     =   5595
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
   ScaleWidth      =   5595
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
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
      Left            =   45
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3045
      Width           =   900
   End
   Begin VB.PictureBox picContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   2250
      ScaleHeight     =   2640
      ScaleWidth      =   3315
      TabIndex        =   2
      Top             =   105
      Width           =   3315
      Begin VB.CheckBox chkRequired 
         Caption         =   "Required"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2100
         TabIndex        =   9
         Top             =   2272
         Width           =   1110
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
         Left            =   0
         TabIndex        =   4
         Top             =   225
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
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   6
         Top             =   945
         Width           =   3270
      End
      Begin VB.TextBox txtDefault 
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
         Left            =   0
         TabIndex        =   8
         Top             =   2220
         Width           =   1890
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
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   1200
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
         Left            =   0
         TabIndex        =   5
         Top             =   720
         Width           =   795
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Default Value"
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
         Left            =   0
         TabIndex        =   7
         Top             =   1995
         Width           =   960
      End
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   10
      Top             =   2925
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   53
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
      Height          =   2310
      IntegralHeight  =   0   'False
      Left            =   45
      Sorted          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   330
      Width           =   1905
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
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
      Left            =   3555
      TabIndex        =   12
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
      Left            =   4575
      TabIndex        =   13
      Top             =   3045
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Parameters"
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
      TabIndex        =   0
      Top             =   105
      Width           =   825
   End
End
Attribute VB_Name = "frmParamMan"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NewParams() As AddInParameter
Dim SelParam As String
Dim SelParamIdx As Integer
Dim IsUpdating As Boolean

Private Sub chkRequired_Click()

    UpdateParams

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    Dim i As Integer
    Dim j As Integer
    
    If lstParams.ListIndex <> -1 Then
        If MsgBox("Are you sure you want to delete the " + SelParam + " parameter?", vbQuestion + vbYesNo, "Delete Confirmation") = vbYes Then
            txtName.Text = vbNullString
            txtDesc.Text = vbNullString
            txtDefault.Text = vbNullString
            lstParams.RemoveItem lstParams.ListIndex
            For i = 1 To UBound(NewParams)
                If NewParams(i).Name = SelParam Then
                    For j = i + 1 To UBound(NewParams) - 1
                        NewParams(j) = NewParams(j + 1)
                    Next j
                    Exit For
                End If
            Next i
            ReDim Preserve NewParams(UBound(NewParams) - 1)
        End If
    End If
    
    SelParamIdx = -1

End Sub

Private Sub cmdOK_Click()

    If (cmdOK.Caption = "&Update") Then
        params = NewParams
    Else
        ReDim Preserve params(UBound(params) + 1)
        With params(UBound(params))
            .Name = txtName.Text
            .Description = txtDesc.Text
            .Default = txtDefault.Text
            .Required = (chkRequired.Value = vbChecked)
        End With
    End If
    Unload Me

End Sub

Private Sub Form_Load()

    Dim i As Integer

    CenterForm Me
    SetupCharset Me
    
    NewParams = params
    
    For i = 1 To UBound(NewParams)
        lstParams.AddItem NewParams(i).Name
    Next i
    SelParamIdx = -1

End Sub

Private Sub lstParams_Click()

    Dim i As Integer
    
    IsUpdating = True
    
    SelParam = lstParams.List(lstParams.ListIndex)
    
    For i = 1 To UBound(NewParams)
        With NewParams(i)
            If .Name = SelParam Then
                SelParamIdx = i
                txtName.Text = .Name
                txtDesc.Text = .Description
                txtDefault.Text = .Default
                chkRequired.Value = Abs(.Required)
            End If
        End With
    Next i
    
    txtName.Enabled = (SelParamIdx <> -1)
    txtDesc.Enabled = (SelParamIdx <> -1)
    txtDefault.Enabled = (SelParamIdx <> -1)
    chkRequired.Enabled = (SelParamIdx <> -1)
    cmdDelete.Enabled = (SelParamIdx <> -1)
    
    IsUpdating = False

End Sub

Private Sub txtDefault_Change()

    UpdateParams

End Sub

Private Sub txtDesc_Change()
    
    UpdateParams

End Sub

Private Sub txtName_Change()

    If (cmdOK.Caption = "&Update") Then
        UpdateParams
        lstParams.List(lstParams.ListIndex) = txtName.Text
    End If

End Sub

Private Sub UpdateParams()

    If IsUpdating Then Exit Sub

    If SelParamIdx <> -1 Then
        With NewParams(SelParamIdx)
            .Name = txtName.Text
            .Description = txtDesc.Text
            .Default = txtDefault.Text
            .Required = (chkRequired.Value = vbChecked)
        End With
    End If

End Sub
