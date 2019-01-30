VERSION 5.00
Begin VB.Form frmUndoRedoStates 
   BorderStyle     =   0  'None
   ClientHeight    =   3405
   ClientLeft      =   5940
   ClientTop       =   5505
   ClientWidth     =   4470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3405
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtMode 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
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
      Left            =   105
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "..."
      Top             =   1920
      Width           =   1560
   End
   Begin VB.ListBox lstStates 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1785
      Left            =   75
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   75
      Width           =   3900
   End
   Begin VB.CommandButton cmdStateOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   2415
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1905
      Width           =   750
   End
   Begin VB.CommandButton cmdStateCancel 
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
      Height          =   300
      Left            =   3225
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1905
      Width           =   750
   End
   Begin VB.CommandButton cmdStateBack 
      Enabled         =   0   'False
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4065
   End
End
Attribute VB_Name = "frmUndoRedoStates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdStateCancel_Click()

    SelRedoUndo = sCancel
    Unload Me

End Sub

Private Sub cmdStateOK_Click()

    SelRedoUndoCount = lstStates.SelCount - 1
    Unload Me

End Sub

Private Sub Form_Load()

    LocalizeUI

    Width = cmdStateBack.Width
    Height = cmdStateBack.Height

End Sub

Private Sub lstStates_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    Dim i As Integer
    
    If lstStates.SelCount Then
        For i = 0 To lstStates.ListCount - 1
            If lstStates.Selected(i) Then
                Exit For
            Else
                lstStates.Selected(i) = True
            End If
        Next i
    End If

End Sub

Private Sub lstStates_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    cmdStateOK.Enabled = True

    Select Case SelRedoUndo
        Case sUndo
            txtMode.Text = GetLocalizedStr(265)
        Case sRedo
            txtMode.Text = GetLocalizedStr(266)
    End Select
    
    txtMode.Text = txtMode.Text + " " & lstStates.SelCount & " " + GetLocalizedStr(498)

End Sub

Private Sub LocalizeUI()

    cmdStateOK.Caption = GetLocalizedStr(186)
    cmdStateCancel.Caption = GetLocalizedStr(187)

End Sub
