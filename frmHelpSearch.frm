VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{AC88F9D2-F91F-4585-B5C3-BC2103206F5D}#1.0#0"; "LINE3D.OCX"
Begin VB.Form frmHelpSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search Help"
   ClientHeight    =   4290
   ClientLeft      =   6825
   ClientTop       =   5085
   ClientWidth     =   4545
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHelpSearch.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar sbInfo 
      Align           =   2  'Align Bottom
      Height          =   270
      Left            =   0
      TabIndex        =   10
      Top             =   4020
      Width           =   4545
      _ExtentX        =   8017
      _ExtentY        =   476
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
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
   Begin MSComctlLib.ListView lvIndex 
      Height          =   2340
      Left            =   30
      TabIndex        =   9
      Top             =   1650
      Width           =   3465
      _ExtentX        =   6112
      _ExtentY        =   4128
      SortKey         =   1
      View            =   3
      LabelEdit       =   1
      SortOrder       =   -1  'True
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "pTitle"
         Text            =   "Page Title"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   "pHits"
         Text            =   "Hits"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.OptionButton opSOptions 
      Caption         =   "Match exact phrase"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   2
      Left            =   255
      TabIndex        =   8
      Top             =   1200
      Width           =   1785
   End
   Begin VB.OptionButton opSOptions 
      Caption         =   "Match all words"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   1
      Left            =   255
      TabIndex        =   7
      Top             =   967
      Width           =   1785
   End
   Begin VB.OptionButton opSOptions 
      Caption         =   "Match any word"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   255
      TabIndex        =   6
      Top             =   735
      Value           =   -1  'True
      Width           =   1785
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "&Open"
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
      Height          =   345
      Left            =   3615
      TabIndex        =   5
      Top             =   1650
      Width           =   870
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3615
      TabIndex        =   4
      Top             =   3645
      Width           =   870
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "&Search"
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
      Height          =   345
      Left            =   3615
      TabIndex        =   3
      Top             =   285
      Width           =   870
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   2
      Top             =   1500
      Width           =   4500
      _ExtentX        =   7938
      _ExtentY        =   53
   End
   Begin VB.TextBox txtKeywords 
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
      Left            =   30
      TabIndex        =   1
      Top             =   315
      Width           =   3465
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Keywords"
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
      Left            =   30
      TabIndex        =   0
      Top             =   90
      Width           =   705
   End
End
Attribute VB_Name = "frmHelpSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdOpen_Click()

    If lvIndex.SelectedItem Is Nothing Then Exit Sub

    ShowHelp lvIndex.SelectedItem.key

End Sub

Private Sub cmdSearch_Click()

    DoSearch

End Sub

Private Sub DoSearch()

    Dim k() As String
    Dim cFile As String
    Dim fCont As String
    Dim mCount As Integer
    Dim f As Integer
    Dim nItm As ListItem
    
    lvIndex.ListItems.Clear
    lvIndex.Sorted = True
    cmdOpen.Enabled = False
        
    If InStr(txtKeywords.Text, " ") = 0 Then
        ReDim k(0)
        k(0) = LCase(txtKeywords.Text)
    Else
        k = Split(LCase(txtKeywords.Text), " ")
    End If
    
    cFile = Dir(AppPath + "Help\*.htm*")
    While cFile <> ""
        Select Case LCase(cFile)
        Case "index.html", "contents.htm"
            'just skip these files...
        Case Else
            fCont = ReadFile(cFile)
            
            If opSOptions(0).Value Then
                mCount = MatchAny(fCont, k)
            ElseIf opSOptions(1).Value Then
                mCount = MatchAll(fCont, k)
            Else
                mCount = MatchPhrase(fCont)
            End If
            
            If mCount > 0 Then
                Set nItm = lvIndex.ListItems.Add(, cFile, PageTitle(fCont))
                nItm.SubItems(1) = Format(mCount, "0000")
            End If
            
            f = f + 1
        End Select
        cFile = Dir
    Wend
    
    If lvIndex.ListItems.Count > 0 Then
        lvIndex_ItemClick lvIndex.ListItems(1)
        lvIndex.SetFocus
        
        lvIndex.Sorted = False
        For Each nItm In lvIndex.ListItems
            nItm.SubItems(1) = CStr(Val(nItm.SubItems(1)))
        Next nItm
    End If
    
    sbInfo.SimpleText = lvIndex.ListItems.Count & " matches total from " & f & " available files"
    
    CoolListView lvIndex

End Sub

Private Function MatchPhrase(fCont As String) As Integer

    Dim p As Integer
    Dim c As Integer

    p = 0
    Do While True
        p = InStr(p + 1, LCase(fCont), LCase(txtKeywords.Text))
        If p = 0 Then Exit Do
        c = c + 1
    Loop
        
    MatchPhrase = c

End Function

Private Function MatchAll(fCont As String, k() As String) As Integer

    Dim i As Integer
    Dim p() As Integer
    Dim c As Integer
    
    ReDim p(UBound(k))
    For i = 0 To UBound(p)
        p(i) = 0
    Next i
    
    Do While True
        For i = 0 To UBound(k)
            p(i) = InStr(p(i) + 1, fCont, k(i))
        Next i
        For i = 0 To UBound(p)
            If p(i) = 0 Then Exit Do
        Next i
        c = c + 1
    Loop
    
ReturnResults:
    MatchAll = c

End Function

Private Function MatchAny(fCont As String, k() As String) As Integer

    Dim i As Integer
    Dim p As Integer
    Dim c As Integer
    
    For i = 0 To UBound(k)
        p = InStr(fCont, k(i))
        While p > 0
            c = c + 1
            p = InStr(p + 1, fCont, k(i))
        Wend
    Next i
    
    MatchAny = c

End Function

Private Function PageTitle(fCont As String) As String

    Dim p1 As Integer
    Dim p2 As Integer
    
    On Error Resume Next
    
    p1 = InStr(LCase(fCont), "<title>") + 7
    p2 = InStr(LCase(fCont), "</title>")
    
    PageTitle = Mid(fCont, p1, p2 - p1)

End Function

Private Function ReadFile(FileName As String) As String

    On Error Resume Next
    Open AppPath + "Help\" + FileName For Input As #ff
        ReadFile = Input$(LOF(ff), ff)
    Close #ff

End Function

Private Sub lvIndex_DblClick()

    cmdOpen_Click

End Sub

Private Sub lvIndex_ItemClick(ByVal Item As MSComctlLib.ListItem)

    Item.EnsureVisible
    Item.Selected = True

    cmdOpen.Enabled = True

End Sub

Private Sub txtKeywords_Change()

    cmdSearch.Enabled = (txtKeywords.Text <> "")

End Sub
