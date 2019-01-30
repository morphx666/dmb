VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmPosOffsetAdvanced 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Advanced Offsets Configuration"
   ClientHeight    =   2970
   ClientLeft      =   6120
   ClientTop       =   6780
   ClientWidth     =   6060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPosOffsetAdvanced.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCompile 
      Caption         =   "Compile"
      Height          =   360
      Left            =   1500
      TabIndex        =   8
      Top             =   2520
      Width           =   1290
   End
   Begin VB.CommandButton cmdPreview 
      Caption         =   "Preview..."
      Height          =   360
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   1290
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&Add..."
      Height          =   360
      Left            =   4680
      TabIndex        =   2
      Top             =   330
      Width           =   1290
   End
   Begin VB.CommandButton cmdEdit 
      Caption         =   "&Edit..."
      Enabled         =   0   'False
      Height          =   360
      Left            =   4680
      TabIndex        =   3
      Top             =   750
      Width           =   1290
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "De&lete"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4680
      TabIndex        =   4
      Top             =   1785
      Width           =   1290
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4680
      TabIndex        =   7
      Top             =   2520
      Width           =   1290
   End
   Begin VB.ListBox lstOffsets 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   330
      Width           =   4485
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   15
      TabIndex        =   5
      Top             =   2340
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   53
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Defined Offsets"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   105
      Width           =   1140
   End
End
Attribute VB_Name = "frmPosOffsetAdvanced"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()

    curOffsetStr = ""
    frmPosOffsetAdvancedAdd.Show vbModal
    
    If LenB(curOffsetStr) <> 0 Then Project.CustomOffsets = Project.CustomOffsets + "`" + curOffsetStr
    LoadCustomOffsets
    
End Sub

Private Sub cmdCompile_Click()

    Me.Enabled = False
    frmMain.ToolsCompile
    Me.Enabled = True

End Sub

Private Sub cmdDelete_Click()

    Dim o() As String
    Dim i As Integer
    
    o = Split(Project.CustomOffsets, "`")
    For i = lstOffsets.ListIndex + 1 To lstOffsets.ListCount - 1
        o(i) = o(i + 1)
    Next i
    ReDim Preserve o(UBound(o) - 1)
    Project.CustomOffsets = Join(o, "`")
    LoadCustomOffsets

End Sub

Private Sub cmdEdit_Click()
    
    Dim o() As String
    Dim SelId As Integer
    
    SelId = lstOffsets.ListIndex + 1
    
    o = Split(Project.CustomOffsets, "`")
    curOffsetStr = o(SelId)
    frmPosOffsetAdvancedAdd.Show vbModal
    
    If LenB(curOffsetStr) = 0 Then Exit Sub
    o(SelId) = curOffsetStr
    Project.CustomOffsets = Join(o, "`")
    LoadCustomOffsets
    lstOffsets.ListIndex = SelId - 1

End Sub

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub cmdPreview_Click()

    frmMain.ShowPreview

End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/menus_offset_cusoff.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    LoadCustomOffsets
    
    cmdPreview.Enabled = frmMain.mnuToolsPreview.Enabled
    cmdCompile.Enabled = frmMain.mnuToolsCompile.Enabled

End Sub

Private Sub LoadCustomOffsets()

    Dim o() As String
    Dim i As Integer
    
    lstOffsets.Clear
    If LenB(Project.CustomOffsets) <> 0 Then
        o = Split(Project.CustomOffsets, "`")
        For i = 1 To UBound(o)
            lstOffsets.AddItem Split(o(i), "@")(0)
        Next i
        lstOffsets.ListIndex = 0
        cmdEdit.Enabled = True
        cmdDelete.Enabled = True
    Else
        cmdEdit.Enabled = False
        cmdDelete.Enabled = False
    End If
    
End Sub
