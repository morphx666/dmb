VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmURLBookmark 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Bookmark"
   ClientHeight    =   4455
   ClientLeft      =   6165
   ClientTop       =   5520
   ClientWidth     =   5025
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmURLBookmark.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstBookmarks 
      Height          =   2595
      Left            =   75
      TabIndex        =   1
      Top             =   345
      Width           =   3405
   End
   Begin VB.TextBox txtBookmark 
      Height          =   285
      Left            =   75
      TabIndex        =   3
      Top             =   3405
      Width           =   3405
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4050
      TabIndex        =   6
      Top             =   4005
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   4005
      Width           =   900
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   3855
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   53
   End
   Begin VB.Label lblABookmarks 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Bookmarks"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   1455
   End
   Begin VB.Label lblTBookmark 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Target Bookmark"
      Height          =   195
      Left            =   75
      TabIndex        =   2
      Top             =   3165
      Width           =   1215
   End
End
Attribute VB_Name = "frmURLBookmark"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()
    
    frmMain.txtURL.Text = GetURL(frmMain.txtURL.Text) + IIf(LenB(txtBookmark.Text) = 0, "", "#") + txtBookmark.Text
    
    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/url_bookmark.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    LocalizeUI

    InitDlg

End Sub

Private Sub InitDlg()

    Dim fn As String
    Dim url As String
    Dim bkm As String
    Dim i As Integer
    
    ForceCommandsLinks2Local

    If IsCommand(frmMain.tvMenus.SelectedItem.key) Then
        Select Case frmMain.tsCmdType.SelectedItem.key
            Case "tsOver"
                fn = MenuCmds(GetID).Actions.onmouseover.url
            Case "tsClick"
                fn = MenuCmds(GetID).Actions.onclick.url
            Case "tsDoubleClick"
                fn = MenuCmds(GetID).Actions.OnDoubleClick.url
        End Select
    Else
        Select Case frmMain.tsCmdType.SelectedItem.key
            Case "tsOver"
                fn = MenuGrps(GetID).Actions.onmouseover.url
            Case "tsClick"
                fn = MenuGrps(GetID).Actions.onclick.url
            Case "tsDoubleClick"
                fn = MenuGrps(GetID).Actions.OnDoubleClick.url
        End Select
    End If
    
    RestoreCommandsLinks
    
    url = GetURL(fn)
    bkm = GetBookmark(fn)
    If FileExists(url) Then GetBookmarks url
    
    If LenB(bkm) <> 0 Then
        If Not Exists(bkm) Then
            lstBookmarks.AddItem bkm
        End If
    End If
    
    For i = 0 To lstBookmarks.ListCount - 1
        If lstBookmarks.List(i) = bkm Then
            lstBookmarks.ListIndex = i
            Exit For
        End If
    Next i
    
    caption = "Bookmark - " + GetFileName(fn)

End Sub

Private Function Exists(fn As String) As Boolean

    Dim i As Integer
    
    For i = 0 To lstBookmarks.ListCount - 1
        If lstBookmarks.List(i) = fn Then
            Exists = True
            Exit Function
        End If
    Next i

End Function

Private Function GetURL(ByVal url As String) As String

    If InStr(url, "#") > 0 Then
        url = Left(url, InStrRev(url, "#") - 1)
    End If
    GetURL = url

End Function

Private Function GetBookmark(ByVal url As String) As String

    If InStr(url, "#") > 0 Then
        url = Mid(url, InStrRev(url, "#") + 1)
    Else
        url = ""
    End If
    GetBookmark = url

End Function

Private Sub GetBookmarks(ByVal FileName As String)

    Dim atags() As String
    Dim i As Integer
    Dim aname As String
    Dim fCode As String
    Dim acode As String
    Dim p As Integer
    
    ReDim atags(0)
    ReDim HotSpots(0)
    
    fCode = LoadFile(FileName)
    
    fCode = Replace(fCode, "<A ", "<a ")
    fCode = Replace(fCode, "</A>", "</a>")
    fCode = Replace(fCode, "<a" + vbCrLf, "<a ")
    fCode = Replace(fCode, "<A" + vbCrLf, "<a ")
    atags = Split(fCode, "<a ")
    
    If UBound(atags) > 0 Then
        For i = 1 To UBound(atags)
            p = InStr(atags(i), ">")
            If p > 0 Then
                acode = Left(atags(i), p - 1)
                aname = GetParamVal(acode, "name")
                If LenB(aname) <> 0 Then
                    lstBookmarks.AddItem aname
                End If
            End If
        Next i
    End If

End Sub

Private Sub lstBookmarks_Click()

    txtBookmark.Text = lstBookmarks.List(lstBookmarks.ListIndex)

End Sub

Private Sub lstBookmarks_DblClick()

    If lstBookmarks.ListIndex >= 0 Then cmdOK_Click

End Sub

Private Sub txtBookmark_GotFocus()

    SelAll txtBookmark

End Sub

Private Sub LocalizeUI()

    lblABookmarks.caption = GetLocalizedStr(891)
    lblTBookmark.caption = GetLocalizedStr(892)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
