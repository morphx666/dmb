VERSION 5.00
Begin VB.Form frmOpenAddIn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Open AddIn"
   ClientHeight    =   3300
   ClientLeft      =   7695
   ClientTop       =   6285
   ClientWidth     =   4230
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpenAddIn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3300
   ScaleWidth      =   4230
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   825
      Left            =   75
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   1980
      Width           =   4050
   End
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
      Left            =   75
      TabIndex        =   3
      Top             =   2850
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
      Left            =   3225
      TabIndex        =   5
      Top             =   2850
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
      Left            =   2250
      TabIndex        =   4
      Top             =   2850
      Width           =   900
   End
   Begin VB.ListBox lstAddIns 
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
      Left            =   75
      TabIndex        =   1
      Top             =   285
      Width           =   4050
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Installed AddIns"
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
      Top             =   60
      Width           =   1170
   End
End
Attribute VB_Name = "frmOpenAddIn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If LITE = 0 Then

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdDelete_Click()

    On Error Resume Next
    
    Dim SelAddIn As String
    
    If lstAddIns.ListIndex <> -1 Then
        SelAddIn = lstAddIns.List(lstAddIns.ListIndex)
    
        If MsgBox("Are you sure you want to delete the " + SelAddIn + " AddIn?", vbQuestion + vbYesNo, "Delete Confirmation") = vbYes Then
            Kill AppPath + "AddIns\" + lstAddIns.List(lstAddIns.ListIndex) + ".ext"
            If FileExists(AppPath + "AddIns\" + lstAddIns.List(lstAddIns.ListIndex) + ".par") Then
                FileExists (AppPath + "AddIns\" + lstAddIns.List(lstAddIns.ListIndex) + ".par")
            End If
            GetAddIns
        End If
    End If
    
End Sub

Private Sub cmdOK_Click()

    If lstAddIns.ListIndex >= 0 Then
        Project.AddIn.Name = lstAddIns.List(lstAddIns.ListIndex)
        Project.AddIn.Description = txtDesc.Text
        With frmAddInEditor
            .GetSections "/default/"
            .GetSections Project.AddIn.Name
            .UpdateTitleBar
        End With
    End If
    
    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    
    GetAddIns
    
End Sub

Private Sub GetAddIns()

    Dim fName As String

    lstAddIns.Clear
    fName = Dir(AppPath + "AddIns\*.ext")
    Do Until LenB(fName) = 0
        lstAddIns.AddItem Left$(fName, InStrRev(fName, ".") - 1)
        fName = Dir
    Loop
    
    If lstAddIns.ListCount Then
        lstAddIns.Selected(0) = True
    End If

End Sub

Private Sub lstAddIns_Click()

    Dim ff As Integer
    Dim sStr As String
    
    On Error Resume Next
    
    txtDesc.Text = vbNullString
    ff = FreeFile
    Open AppPath + "AddIns\" + lstAddIns.List(lstAddIns.ListIndex) + ".ext" For Input As #ff
        Line Input #ff, sStr
        Do Until sStr = "***" Or EOF(ff)
            txtDesc.Text = txtDesc.Text + sStr + vbCrLf
            Line Input #ff, sStr
        Loop
        If LenB(txtDesc.Text) <> 0 Then txtDesc.Text = Left$(txtDesc.Text, Len(txtDesc) - 2)
    Close #ff

End Sub

Private Sub lstAddIns_DblClick()

    cmdOK_Click

End Sub

#End If
