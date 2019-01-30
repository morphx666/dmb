VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmAddBrowser 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Browser"
   ClientHeight    =   2250
   ClientLeft      =   5910
   ClientTop       =   6045
   ClientWidth     =   4770
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
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2250
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2715
      TabIndex        =   6
      Top             =   1770
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3780
      TabIndex        =   7
      Top             =   1770
      Width           =   900
   End
   Begin VB.TextBox txtFileName 
      Height          =   315
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1050
      Width           =   4020
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   105
      TabIndex        =   1
      Top             =   375
      Width           =   2460
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   1185
      Top             =   1605
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin SmartButtonProject.SmartButton cmdBrowse 
      Height          =   360
      Left            =   4215
      TabIndex        =   5
      Top             =   1020
      Width           =   420
      _ExtentX        =   741
      _ExtentY        =   635
      Picture         =   "frmAddBrowser.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   2715
      TabIndex        =   2
      Top             =   435
      Width           =   45
   End
   Begin VB.Label lblCommand 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Command"
      Height          =   195
      Left            =   105
      TabIndex        =   3
      Top             =   825
      Width           =   705
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   405
   End
End
Attribute VB_Name = "frmAddBrowser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()

    On Error GoTo ExitSub
    
    With cDlg
        .DialogTitle = GetLocalizedStr(414)
        .filter = GetLocalizedStr(411) + "|*.exe"
        .CancelError = True
        .InitDir = GetFilePath(txtFileName.Text)
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .ShowOpen
        txtFileName.Text = .FileName
    End With
    
ExitSub:
    Exit Sub

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim i As Integer
    
    If LenB(txtName.Text) = 0 Then
        MsgBox GetLocalizedStr(412), vbInformation + vbOKOnly, GetLocalizedStr(410)
        Exit Sub
    End If
    
    If Not FileExists(txtFileName.Text) Then
        MsgBox GetLocalizedStr(413), vbInformation + vbOKOnly, GetLocalizedStr(410)
        Exit Sub
    End If
        
    If cmdOK.Caption = GetLocalizedStr(426) Then
        i = frmBrowsers.lvBrowsers.SelectedItem.Index - 1
    Else
        i = 1
        Do Until LenB(GetSetting(App.EXEName, "Browsers", "Name" & i, vbNullString)) = 0
            i = i + 1
        Loop
    End If
    
    SaveSetting App.EXEName, "Browsers", "Name" & i, txtName.Text
    SaveSetting App.EXEName, "Browsers", "Command" & i, txtFileName.Text
    
    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    LocalizeUI

End Sub

Private Sub LocalizeUI()

    Caption = GetLocalizedStr(410)

    lblName.Caption = GetLocalizedStr(409)
    lblCommand.Caption = GetLocalizedStr(271)

    cmdOK.Caption = GetLocalizedStr(186)
    cmdCancel.Caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub

Private Sub txtFileName_Change()

    lblVersion.Caption = GetFileVersion(Long2Short(txtFileName.Text))

End Sub
