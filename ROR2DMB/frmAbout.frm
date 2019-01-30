VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Object = "{2A2AD7CA-AC77-46F3-84DC-115021432312}#1.0#0"; "hRef.ocx"
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3555
   ClientLeft      =   7365
   ClientTop       =   7095
   ClientWidth     =   6090
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2453.724
   ScaleMode       =   0  'User
   ScaleWidth      =   5718.826
   ShowInTaskbar   =   0   'False
   Begin href.uchref1 rorLink 
      Height          =   315
      Left            =   1380
      TabIndex        =   9
      ToolTipText     =   "Click to visit the official ROR file format web site"
      Top             =   2310
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   556
      Caption         =   "Click here for more information about the ROR format"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
      ForeColor       =   16711680
      URL             =   "http://www.rorweb.com"
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   -15
      TabIndex        =   5
      Top             =   2925
      Width           =   6090
      _ExtentX        =   10742
      _ExtentY        =   53
   End
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   90
      Picture         =   "frmAbout.frx":038A
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   90
      Width           =   480
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4710
      TabIndex        =   0
      Top             =   3105
      Width           =   1260
   End
   Begin VB.Label lblBasedOn 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Based on..."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   165
      Left            =   615
      TabIndex        =   10
      Top             =   435
      Width           =   690
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "XML Parser based on the ROR format:"
      Height          =   195
      Left            =   255
      TabIndex        =   8
      Top             =   1440
      Width           =   2745
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   """So Search Engines Can Understand Your Website!"""
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1380
      TabIndex        =   7
      Top             =   2025
      Width           =   3720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ROR (Resources of a Resource)"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   1395
      TabIndex        =   6
      Top             =   1800
      Width           =   3075
   End
   Begin VB.Image imgRORLogo 
      Height          =   1485
      Left            =   30
      Picture         =   "frmAbout.frx":2D0C
      Top             =   1665
      Width           =   1485
   End
   Begin VB.Label lblCopyRight 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "cpy"
      Height          =   195
      Left            =   255
      TabIndex        =   4
      Top             =   1020
      Width           =   255
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Application Title"
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000009&
      Height          =   360
      Left            =   630
      TabIndex        =   2
      Top             =   45
      Width           =   2250
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   195
      Left            =   255
      TabIndex        =   3
      Top             =   780
      Width           =   525
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000006&
      FillColor       =   &H80000002&
      FillStyle       =   0  'Solid
      Height          =   690
      Left            =   -15
      Top             =   -15
      Width           =   6120
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Unload Me
    
End Sub

Private Sub Form_Load()

    CenterForm Me

    Me.Caption = "About " & frmMain.Caption
    lblVersion.Caption = "Version " & App.Major & "." & App.Minor & "." & App.Revision
    lblTitle.Caption = frmMain.Caption
    lblCopyRight.Caption = "© " & Year(Now) & " xFX JumpStart"
    
    lblBasedOn.Caption = "Based on DHTML Menu Builder " + GetFileVersion(AppPath + "dmb.exe")
    
End Sub
