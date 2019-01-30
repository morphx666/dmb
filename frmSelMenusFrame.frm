VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmSelMenusFrame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Menus Frame"
   ClientHeight    =   2190
   ClientLeft      =   6975
   ClientTop       =   6705
   ClientWidth     =   4365
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
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   3600
      Top             =   180
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   60
      TabIndex        =   2
      Top             =   1680
      Width           =   4260
      _ExtentX        =   7514
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3420
      TabIndex        =   3
      Top             =   1770
      Width           =   900
   End
   Begin VB.ListBox lstFrames 
      Height          =   840
      Left            =   60
      TabIndex        =   1
      Top             =   720
      Width           =   4260
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmSelMenusFrame.frx":0000
      Height          =   615
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   4170
   End
End
Attribute VB_Name = "frmSelMenusFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    MenusFrame = lstFrames.ItemData(lstFrames.ListIndex)
    Unload Me

End Sub

Private Sub Form_Load()

    SetupCharset Me
    LocalizeUI
    
    tmrInit.Enabled = True

End Sub

Private Sub tmrInit_Timer()
    
    Dim i As Integer
    Dim t As Integer
    
    tmrInit.Enabled = False

    CenterForm Me
    
    On Error Resume Next
    
    t = UBound(FramesInfo.Frames)
    For i = 1 To t
        If LCase(FramesInfo.Frames(i).srcFile) <> LCase(GetRealLocal.HotSpotEditor.HotSpotsFile) Then
            lstFrames.AddItem FramesInfo.Frames(i).Name
            lstFrames.ItemData(lstFrames.NewIndex) = i
        End If
    Next i
    
    If lstFrames.ListCount > 0 Then
        lstFrames.ListIndex = 0
        
        If lstFrames.ListCount = 1 Then
            cmdOK_Click
        End If
    Else
        
        cmdOK.Enabled = False
        Unload Me
    End If

End Sub

Private Sub LocalizeUI()

    lblInfo.Caption = GetLocalizedStr(893)
    
    cmdOK.Caption = GetLocalizedStr(186)
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
    End If

End Sub
