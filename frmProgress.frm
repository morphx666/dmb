VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmProgress 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1485
   ClientLeft      =   6960
   ClientTop       =   5775
   ClientWidth     =   4770
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
   ScaleHeight     =   1485
   ScaleWidth      =   4770
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ProgressBar pbGroups 
      Height          =   240
      Left            =   195
      TabIndex        =   2
      Top             =   375
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   375
      Left            =   1927
      TabIndex        =   0
      Top             =   1590
      Width           =   900
   End
   Begin MSComctlLib.ProgressBar pbCommands 
      Height          =   240
      Left            =   195
      TabIndex        =   4
      Top             =   1050
      Width           =   4365
      _ExtentX        =   7699
      _ExtentY        =   423
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.Image Image2 
      Height          =   240
      Left            =   135
      Picture         =   "frmProgress.frx":0000
      Top             =   105
      Width           =   240
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   135
      Picture         =   "frmProgress.frx":0102
      Top             =   795
      Width           =   240
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Menu Commands"
      Height          =   210
      Left            =   435
      TabIndex        =   3
      Top             =   810
      Width           =   1395
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Groups"
      Height          =   210
      Left            =   435
      TabIndex        =   1
      Top             =   120
      Width           =   570
   End
End
Attribute VB_Name = "frmProgress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()

    CenterForm Me
    
    frmMain.tvMenus.Enabled = False
    frmMain.Enabled = False
    Busy = True
    UserCanceled = False
    
    pbCommands.Value = 0
    pbCommands.Min = 0
    pbGroups.Min = 0
    pbGroups.Value = 0
    
End Sub

Public Sub CloseDialog()

    Busy = False
    
    frmMain.Enabled = True
    frmMain.tvMenus.Enabled = True

    Unload Me

End Sub

