VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmHSArrangement 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Hotspots Arrangement"
   ClientHeight    =   1935
   ClientLeft      =   7155
   ClientTop       =   6645
   ClientWidth     =   4185
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
   Icon            =   "frmHSArrangement.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1125
      Top             =   1500
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3225
      TabIndex        =   6
      Top             =   1500
      Width           =   900
   End
   Begin VB.OptionButton opAlignmentStyle 
      Height          =   270
      Index           =   0
      Left            =   600
      Picture         =   "frmHSArrangement.frx":038A
      Style           =   1  'Graphical
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   495
      Value           =   -1  'True
      Width           =   345
   End
   Begin VB.OptionButton opAlignmentStyle 
      Height          =   270
      Index           =   1
      Left            =   600
      Picture         =   "frmHSArrangement.frx":04D4
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   795
      Width           =   345
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   0
      TabIndex        =   5
      Top             =   1350
      Width           =   4185
      _ExtentX        =   7382
      _ExtentY        =   53
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Select how the items on your menu bar are arranged:"
      Height          =   405
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   3990
   End
   Begin VB.Label lblASVertical 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Vertically"
      Height          =   195
      Left            =   1035
      TabIndex        =   2
      Top             =   525
      Width           =   645
   End
   Begin VB.Label lblASHorizontal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Horizontally"
      Height          =   195
      Left            =   1035
      TabIndex        =   4
      Top             =   825
      Width           =   840
   End
End
Attribute VB_Name = "frmHSArrangement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

    Dim g As Integer
    
    For g = 1 To UBound(MenuGrps)
        If opAlignmentStyle(0).Value Then
            MenuGrps(g).x = 0
            MenuGrps(g).y = g
        Else
            MenuGrps(g).x = g
            MenuGrps(g).y = 0
        End If
    Next g
    
    Unload Me

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    LocalizeUI
    
    tmrInit.Enabled = True

End Sub

Private Sub LocalizeUI()

    cmdOK.Caption = GetLocalizedStr(186)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
    End If

End Sub

Private Sub opAlignmentStyle_Click(Index As Integer)

    cmdOK.SetFocus

End Sub

Private Sub tmrInit_Timer()

    Dim i As Integer
    Dim j As Integer
    
    tmrInit.Enabled = False
    
    For i = 1 To UBound(MenuGrps)
        If MenuGrps(i).x <> MenuGrps(i).y Then
            For j = i + 1 To UBound(MenuGrps)
                If MenuGrps(j).x <> MenuGrps(j).y Then
                    If MenuGrps(i).x = MenuGrps(j).x Then
                        opAlignmentStyle(0).Value = True
                    Else
                        opAlignmentStyle(1).Value = True
                    End If
                    Exit For
                End If
            Next j
            Exit For
        End If
    Next i

End Sub
