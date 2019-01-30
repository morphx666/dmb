VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmEditLink 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Link"
   ClientHeight    =   1380
   ClientLeft      =   6750
   ClientTop       =   6945
   ClientWidth     =   4785
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmEditLink.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1380
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3690
      TabIndex        =   4
      Top             =   900
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2595
      TabIndex        =   3
      Top             =   900
      Width           =   975
   End
   Begin VB.TextBox txtLink 
      Height          =   285
      Left            =   135
      TabIndex        =   1
      Top             =   360
      Width           =   4140
   End
   Begin SmartButtonProject.SmartButton cmdBrowse 
      Height          =   315
      Left            =   4305
      TabIndex        =   2
      Top             =   345
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmEditLink.frx":058A
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
   Begin VB.Label lblItemName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "item name"
      Height          =   195
      Left            =   135
      TabIndex        =   0
      Top             =   150
      Width           =   735
   End
End
Attribute VB_Name = "frmEditLink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowse_Click()

    frmMain.SelectItem frmMain.tvMenus.Nodes(frmBrokenLinksReport.lvLinks.SelectedItem.Tag), True
    frmMain.cmdBrowse_Click
    
    txtLink.Text = frmMain.txtURL.Text

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    frmBrokenLinksReport.lvLinks.SelectedItem.SubItems(1) = txtLink.Text
    Unload Me

End Sub

Private Sub Form_Load()

    SetupCharset Me
    CenterForm Me
    LocalizeUI

End Sub

Private Sub LocalizeUI()

    Caption = GetLocalizedStr(938)
    
    cmdOK.Caption = GetLocalizedStr(186)
    cmdCancel.Caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
