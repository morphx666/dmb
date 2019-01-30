VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmFontSubst 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font Substitution"
   ClientHeight    =   3360
   ClientLeft      =   6870
   ClientTop       =   7050
   ClientWidth     =   5325
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFontSubst.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ListView lvFonts 
      Height          =   1695
      Left            =   105
      TabIndex        =   1
      Top             =   345
      Width           =   3900
      _ExtentX        =   6879
      _ExtentY        =   2990
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      HideColumnHeaders=   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   5
      Top             =   2805
      Width           =   5265
      _ExtentX        =   9287
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Height          =   360
      Left            =   2805
      TabIndex        =   6
      Top             =   2940
      Width           =   1155
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add..."
      Enabled         =   0   'False
      Height          =   360
      Left            =   4125
      TabIndex        =   4
      Top             =   2332
      Width           =   1155
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   360
      Left            =   4125
      TabIndex        =   7
      Top             =   2940
      Width           =   1155
   End
   Begin VB.TextBox txtSubst 
      Enabled         =   0   'False
      Height          =   285
      Left            =   105
      TabIndex        =   3
      Top             =   2370
      Width           =   3900
   End
   Begin VB.Label lblSubstitutions 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Substitutions"
      Height          =   195
      Left            =   105
      TabIndex        =   2
      Top             =   2145
      Width           =   930
   End
   Begin VB.Label lblFontsList 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fonts used in this project"
      Height          =   195
      Left            =   105
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frmFontSubst"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fs() As String

Private Sub cmdAdd_Click()

    With SelFont
        .Name = lvFonts.SelectedItem.Text
        .Size = 13
        .IsSubst = True
        frmFontDialog.Show vbModal
        
        If .IsValid Then
            AddSubst .Name
        End If
    End With

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim fItem As ListItem
    Dim FontSubsts As String
    
    For Each fItem In lvFonts.ListItems
        If LenB(fItem.key) <> 0 Then FontSubsts = FontSubsts + "|" + fItem.Text + "|" + fItem.key
    Next fItem
    FontSubsts = Mid(FontSubsts, 2)
    
    Project.FontSubstitutions = FontSubsts
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/font_subst.htm"

End Sub

Private Sub Form_Load()

    Dim i As Integer
    
    LocalizeUI
    CenterForm Me
    
    If InStr(Project.FontSubstitutions, "|") > 0 Then
        fs = Split(Project.FontSubstitutions, "|")
    Else
        ReDim fs(0)
    End If
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            If .Name <> "[SEP]" Then
                AddFont .NormalFont.FontName
                AddFont .HoverFont.FontName
            End If
        End With
    Next i
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            AddFont .DefHoverFont.FontName
            AddFont .DefNormalFont.FontName
        End With
    Next i
    
    If lvFonts.ListItems.Count > 0 Then
        With lvFonts.ListItems(1)
            .Selected = True
            .EnsureVisible
        End With
        lvFonts_Click
    End If
    
    CoolListView lvFonts

End Sub

Private Sub AddSubst(FontName As String)

    Dim s() As String
    Dim i As Integer
    
    If InStr(txtSubst.Text, ",") Then
        s = Split(txtSubst.Text, ",")
    Else
        ReDim s(0)
        s(0) = txtSubst.Text
    End If
    
    For i = 0 To UBound(s)
        If s(i) = FontName Then Exit Sub
    Next i
    
    txtSubst.Text = txtSubst.Text + IIf(LenB(txtSubst.Text) = 0, "", ", ") + FontName
    txtSubst.SelStart = Len(txtSubst.Text)
    txtSubst_DblClick

End Sub

Private Sub AddFont(FontName As String)

    Dim i As Integer
    Dim nItem As ListItem
    
    On Error Resume Next
    
    If lvFonts.FindItem(FontName, lvwText, , lvwWhole) Is Nothing Then
        Set nItem = lvFonts.ListItems.Add(, , FontName)
        
        For i = 0 To UBound(fs) / 2
            If fs(i * 2) = FontName Then nItem.key = fs(i * 2 + 1)
        Next i
    End If

End Sub

Private Sub lvFonts_Click()

    txtSubst.Enabled = Not (lvFonts.SelectedItem Is Nothing)
    cmdAdd.Enabled = txtSubst.Enabled
    
    If txtSubst.Enabled Then
        txtSubst.Text = lvFonts.SelectedItem.key
    End If

End Sub

Private Sub txtSubst_Change()

    On Error Resume Next
    lvFonts.SelectedItem.key = txtSubst.Text

End Sub

Private Sub txtSubst_DblClick()

    Dim s As Integer
    Dim e As Integer
    
    If LenB(txtSubst.Text) <> 0 Then
    
        For s = txtSubst.SelStart To 1 Step -1
            If Mid(txtSubst.Text, s, 1) = "," Then Exit For
        Next s
        s = s + 1
        If Mid(txtSubst.Text, s, 1) = " " Then s = s + 1
        e = InStr(s, txtSubst.Text, ",")
        If e = 0 Then e = Len(txtSubst.Text)
        If InStr(e, txtSubst.Text, " ") > e Then e = InStr(e, txtSubst.Text, " ")
        
        txtSubst.SelStart = s - 1
        txtSubst.SelLength = (e - s) + 1
    
        txtSubst.SetFocus
    End If

End Sub

Private Sub LocalizeUI()

    lblFontsList.caption = GetLocalizedStr(714)
    lblSubstitutions.caption = GetLocalizedStr(715)
    
    cmdAdd.caption = GetLocalizedStr(338)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdClose.caption = GetLocalizedStr(424)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdClose.Width = SetCtrlWidth(cmdClose)
    End If
    
End Sub
