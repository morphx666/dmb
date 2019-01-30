VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmImportFromTXT 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import from Text File"
   ClientHeight    =   5280
   ClientLeft      =   12120
   ClientTop       =   5850
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4095
      TabIndex        =   10
      Top             =   4770
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Start"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   375
      Left            =   3030
      TabIndex        =   9
      Top             =   4770
      Width           =   900
   End
   Begin MSComctlLib.ListView lvPreview 
      Height          =   2415
      Left            =   135
      TabIndex        =   8
      Top             =   2190
      Width           =   4830
      _ExtentX        =   8520
      _ExtentY        =   4260
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "Options"
      Height          =   1305
      Left            =   135
      TabIndex        =   3
      Top             =   720
      Width           =   4830
      Begin VB.TextBox txtIgnore 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1380
         TabIndex        =   7
         Text            =   "1"
         Top             =   840
         Width           =   600
      End
      Begin VB.CheckBox chkIgnore 
         Caption         =   "Ignore First                   Rows"
         Height          =   195
         Left            =   255
         TabIndex        =   6
         Top             =   885
         Width           =   2730
      End
      Begin VB.ComboBox cmbDLC 
         Height          =   315
         ItemData        =   "frmImportFromTXT.frx":0000
         Left            =   945
         List            =   "frmImportFromTXT.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   375
         Width           =   1035
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Delimiter"
         Height          =   195
         Left            =   255
         TabIndex        =   5
         Top             =   435
         Width           =   600
      End
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   135
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   4335
   End
   Begin SmartButtonProject.SmartButton cmdBrowse 
      Height          =   315
      Left            =   4605
      TabIndex        =   1
      Top             =   345
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmImportFromTXT.frx":0004
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
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   465
      Top             =   4695
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "0"
      DialogTitle     =   "Select Project"
      Filter          =   "DHTML Menu Builder Projects|*.dmb"
      MaxFileSize     =   1024
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Source Text File"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   120
      Width           =   1155
   End
End
Attribute VB_Name = "frmImportFromTXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkIgnore_Click()

    LoadTXTFile

End Sub

Private Sub cmbDLC_Click()

    LoadTXTFile

End Sub

Private Sub cmdBrowse_Click()

    With cDlg
        .DialogTitle = "Select Source Text File"
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNNoReadOnlyReturn
        .filter = "Text Files|*.txt|CSV Files|*.csv|All Files|*.*"
        .FileName = ""
        On Error Resume Next
        .ShowOpen
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then
            Exit Sub
        End If
        txtFileName.Text = .FileName
        
        LoadTXTFile
    End With

End Sub

Private Sub LoadTXTFile()

    Dim data As String
    Dim lines() As String
    Dim eol As String
    Dim numCols As Integer
    Dim del As String
    
    lvPreview.ListItems.Clear
    lvPreview.ColumnHeaders.Clear
    
    cmdOK.Enabled = False
    If Not FileExists(txtFileName.Text) Then Exit Sub
    
    data = LoadFile(txtFileName.Text)
    If InStr(data, vbCrLf) Then
        eol = vbCrLf
    ElseIf InStr(data, vbCr) Then
        eol = vbCr
    ElseIf InStr(data, vbLf) Then
        eol = vbLf
    End If
    
    lines = Split(data, eol)
    
    Select Case cmbDLC.Text
        Case "[SPACE]"
            del = " "
        Case "[TAB]"
            del = Chr(9)
        Case Else
            del = cmbDLC.Text
    End Select
    
    Dim i As Integer
    Dim k As Integer
    Dim nc As Integer
    Dim start As Integer
    
    If chkIgnore.Value = vbChecked Then
        start = Val(txtIgnore.Text)
    Else
        start = 0
    End If
    
    For i = start To UBound(lines) - 1
        nc = UBound(Split(lines(i), del))
        If nc > numCols Then numCols = nc
    Next i
    
    For i = 1 To numCols + 1
        Select Case i
            Case 1
                lvPreview.ColumnHeaders.Add , , "Parent"
            Case 2
                lvPreview.ColumnHeaders.Add , , "Item"
            Case 3
                lvPreview.ColumnHeaders.Add , , "URL"
            Case Else
                lvPreview.ColumnHeaders.Add , , "Ignored " + CStr(i)
        End Select
    Next i
    
    Dim items() As String
    Dim item As ListItem
    For i = start To UBound(lines) - 1
        items = Split(lines(i), del)
        Set item = lvPreview.ListItems.Add(, "txt" + CStr(i), items(0))
        item.Tag = FixItemName("txt" + item.Text)
        For k = 1 To UBound(items)
            With item.ListSubItems.Add(, item.key + CStr(k), items(k))
                .Tag = FixItemName("txt" + .Text)
            End With
        Next k
    Next i
    
    cmdOK.Enabled = lvPreview.ListItems.Count > 0
    
    CoolListView lvPreview

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

          Dim item As ListItem
          Dim subItem As ListSubItem
          Dim lastSubKey As String
          Dim parent As String
          
   On Error GoTo cmdOK_Click_Error

10        Me.Enabled = False
          
          Dim masterParent As Node
20        Set masterParent = frmMain.tvMapView.SelectedItem
          
30        For Each item In lvPreview.ListItems
40            If item.Text = "" Then
50                If IsTBMapSel Or UBound(MenuGrps) = 0 Then
60                    parent = ""
70                Else
80                    parent = masterParent.Text
90                End If
100           Else
110               parent = item.Text
120           End If

130           AddItem parent, item.SubItems(1)
140       Next
          
150       frmMain.RefreshMap
          
160       Unload Me

   On Error GoTo 0
   Exit Sub

cmdOK_Click_Error:

    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure ParseTextFile at line " & Erl & ". The text file may not be in the correct format."
          
End Sub

Private Sub AddItem(parent As String, item As String)
    Dim id As Integer
       
    id = GetIDFromCaption(parent)
    
    If IsGroupFromCaption(parent) Or parent = "" Then
        If parent = "" Then
            With TemplateGroup
                .Name = "tbg" + FixItemName(item)
                .caption = item
            End With
            AddMenuGroup GetGrpParams(TemplateGroup), True
            
            id = UBound(MenuGrps)
            
            With Project.Toolbars(1)
                ReDim Preserve .Groups(UBound(.Groups) + 1)
                .Groups(UBound(.Groups)) = MenuGrps(UBound(MenuGrps)).Name
            End With
            
            Exit Sub
        End If
        
        If MenuGrps(id).Actions.onmouseover.Type <> atcCascade Then
            With MenuGrps(id).Actions.onmouseover
                .Type = atcCascade
                .TargetMenu = id
            End With
        End If
        
        With TemplateCommand
            .caption = item
            .Name = "ctxt" + FixItemName(item)
            .disabled = False
            .parent = id
        End With
        AddMenuCommand GetCmdParams(TemplateCommand), True, True, True
    Else
        If MenuCmds(id).Actions.onmouseover.Type <> atcCascade Then
            With TemplateGroup
                .Name = "g" + FixItemName(MenuCmds(id).Name)
                .caption = .Name
            End With
            AddMenuGroup GetGrpParams(TemplateGroup), True
            
            With MenuCmds(id).Actions.onmouseover
                .Type = atcCascade
                .TargetMenu = UBound(MenuGrps)
                .TargetMenuAlignment = gacRightTop
            End With
        End If
        
        parent = MenuGrps(MenuCmds(id).Actions.onmouseover.TargetMenu).caption
        AddItem parent, item
    End If
    
End Sub

Private Sub Form_Load()

    CenterForm Me

    cmbDLC.AddItem ","
    cmbDLC.AddItem "."
    cmbDLC.AddItem "+"
    cmbDLC.AddItem "-"
    cmbDLC.AddItem "|"
    cmbDLC.AddItem "/"
    cmbDLC.AddItem "\"
    cmbDLC.AddItem "[TAB]"
    cmbDLC.AddItem "[SPACE]"
    cmbDLC.ListIndex = 0

End Sub

Private Sub txtIgnore_Change()

    If chkIgnore.Value = vbChecked Then LoadTXTFile

End Sub

Private Function GetIDFromCaption(caption As String) As Integer

    Dim i As Integer

    For i = 1 To UBound(MenuGrps)
        If MenuGrps(i).caption = caption Then
            GetIDFromCaption = i
            Exit Function
        End If
    Next i
    
    For i = 1 To UBound(MenuCmds)
        If MenuCmds(i).caption = caption Then
            GetIDFromCaption = i
            Exit Function
        End If
    Next i
    
    GetIDFromCaption = 0

End Function


Private Function IsGroupFromCaption(caption As String) As Boolean

    Dim i As Integer

    For i = 1 To UBound(MenuGrps)
        If MenuGrps(i).caption = caption Then
            IsGroupFromCaption = True
            Exit Function
        End If
    Next i
    
    IsGroupFromCaption = False

End Function

