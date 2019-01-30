VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmPosOffset 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Position Offset"
   ClientHeight    =   3660
   ClientLeft      =   6540
   ClientTop       =   5085
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFramesPosOffset.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3660
   ScaleWidth      =   4695
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAdvanced 
      Caption         =   "Advanced..."
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
      Left            =   60
      TabIndex        =   14
      Top             =   3225
      Width           =   1110
   End
   Begin xfxLine3D.ucLine3D uc3DLine2 
      Height          =   30
      Left            =   75
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1590
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   53
   End
   Begin VB.TextBox txtHS 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   7
      Text            =   "888"
      Top             =   2265
      Width           =   480
   End
   Begin VB.TextBox txtVS 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   9
      Text            =   "888"
      Top             =   2625
      Width           =   480
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
      Left            =   3750
      TabIndex        =   13
      Top             =   3225
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
      Left            =   2730
      TabIndex        =   12
      Top             =   3225
      Width           =   900
   End
   Begin VB.TextBox txtV 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   3
      Text            =   "888"
      Top             =   1140
      Width           =   480
   End
   Begin VB.TextBox txtH 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   765
      TabIndex        =   1
      Text            =   "-888"
      Top             =   780
      Width           =   480
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3075
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   53
   End
   Begin VB.Label lblRootMsg 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmFramesPosOffset.frx":038A
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   75
      TabIndex        =   0
      Top             =   75
      Width           =   4545
   End
   Begin VB.Label lblSubMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Select the vertical and horizonal offset you want to apply to the default position of the submenus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   75
      TabIndex        =   6
      Top             =   1755
      Width           =   4545
   End
   Begin VB.Label lblSubH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixels Horizontally"
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
      Left            =   1320
      TabIndex        =   8
      Top             =   2310
      Width           =   1290
   End
   Begin VB.Label lblSubV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixels Vertically"
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
      Left            =   1320
      TabIndex        =   10
      Top             =   2670
      Width           =   1095
   End
   Begin VB.Label lblRootV 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixels Vertically"
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
      Left            =   1320
      TabIndex        =   4
      Top             =   1185
      Width           =   1095
   End
   Begin VB.Label lblRootH 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Pixels Horizontally"
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
      Left            =   1320
      TabIndex        =   2
      Top             =   825
      Width           =   1290
   End
End
Attribute VB_Name = "frmPosOffset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdvanced_Click()

    frmPosOffsetAdvanced.Show vbModal

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    With Project.MenusOffset
        .RootMenusX = Val(txtH.Text)
        .RootMenusY = Val(txtV.Text)
        
        .SubMenusX = Val(txtHS.Text)
        .SubMenusY = Val(txtVS.Text)
    End With
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/menus_offset.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    LocalizeUI

    With Project.MenusOffset
        txtH.Text = .RootMenusX
        txtV.Text = .RootMenusY
        
        txtHS.Text = .SubMenusX
        txtVS.Text = .SubMenusY
    End With

End Sub

Private Sub txtH_GotFocus()

    SelAll txtH

End Sub

Private Sub txtH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtHS_GotFocus()

    SelAll txtHS

End Sub

Private Sub txtHS_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtV_GotFocus()

    SelAll txtV

End Sub

Private Sub txtV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtVS_GotFocus()

    SelAll txtVS

End Sub

Private Sub txtVS_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 45 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub LocalizeUI()

    lblRootMsg.caption = GetLocalizedStr(432)
    lblSubMsg.caption = GetLocalizedStr(433)
    
    lblRootH.caption = GetLocalizedStr(363)
    lblRootV.caption = GetLocalizedStr(364)
    
    lblSubH.caption = GetLocalizedStr(363)
    lblSubV.caption = GetLocalizedStr(364)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
