VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DBF30C82-CAF3-11D5-84FF-0050BA3D926D}#8.5#0"; "VLMnuPlus.ocx"
Begin VB.Form frmAddInEditor 
   Caption         =   "AddIn Editor - [Untitled]"
   ClientHeight    =   4770
   ClientLeft      =   2970
   ClientTop       =   4590
   ClientWidth     =   7305
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAddInEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   7305
   Begin VB.TextBox txtCodeEditor 
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   3405
      HideSelection   =   0   'False
      Index           =   1
      Left            =   3630
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   720
      Width           =   3450
   End
   Begin VB.TextBox txtCodeEditor 
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3300
      HideSelection   =   0   'False
      Index           =   0
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   3
      Top             =   765
      Width           =   3450
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   3030
      Top             =   4185
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer tmrInit 
      Interval        =   50
      Left            =   3945
      Top             =   4245
   End
   Begin MSComctlLib.ImageList ilStatus 
      Left            =   4455
      Top             =   4170
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddInEditor.frx":038A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddInEditor.frx":07DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddInEditor.frx":0C3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmAddInEditor.frx":109E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtParams 
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   3645
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   315
      Width           =   3480
   End
   Begin VB.TextBox txtParams 
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   3480
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H80000000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   690
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4080
      Width           =   7290
   End
   Begin MSComctlLib.ImageCombo cmbSection 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   30
      Width           =   3480
      _ExtentX        =   6138
      _ExtentY        =   609
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      ImageList       =   "ilStatus"
   End
   Begin VLMnuPlus.VLMenuPlus vlmCtrl 
      Left            =   3525
      Top             =   30
      _ExtentX        =   847
      _ExtentY        =   847
      _CXY            =   4
      _CGUID          =   43165.2824652778
      Language        =   0
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu mnuFileOpenSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuFileOpenSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuFileSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileProperties 
         Caption         =   "&Properties..."
      End
      Begin VB.Menu mnuFileSep04 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuCode 
      Caption         =   "&Function"
      Begin VB.Menu mnuCodeAdd 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mnuCodeSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodeEdit 
         Caption         =   "&Edit..."
      End
      Begin VB.Menu mnuCodeRemove 
         Caption         =   "Re&move"
      End
      Begin VB.Menu mnuCodeSep02 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodeRestore 
         Caption         =   "&Restore"
      End
      Begin VB.Menu mnuCodeSep03 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCodeViewOriginal 
         Caption         =   "&View Original"
      End
   End
   Begin VB.Menu mnuParam 
      Caption         =   "Parameter"
      Begin VB.Menu mnuParamAdd 
         Caption         =   "&Add..."
      End
      Begin VB.Menu mnuParamEdit 
         Caption         =   "&Edit..."
      End
      Begin VB.Menu mnuParamSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuParamInsert 
         Caption         =   "&Insert..."
      End
   End
   Begin VB.Menu mnuOp 
      Caption         =   "&Options"
      Begin VB.Menu mnuOpFont 
         Caption         =   "&Font..."
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "Contents..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpSep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About..."
      End
   End
End
Attribute VB_Name = "frmAddInEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If LITE = 0 Then

Dim ff As Integer
Dim OriginalAddIn As AddInDef
Dim SecBackup() As Section
Dim ParBackup() As AddInParameter
Dim PropChanged As Boolean
Dim pFont As SelFontDef
Dim csChangedFlag As Boolean
Dim csIsOldAddIn As Boolean
Dim xMenu As CMenu

Private Function HasChanged() As Boolean

    Dim i As Integer
    
    HasChanged = (UBound(Sections) <> UBound(SecBackup)) Or PropChanged
    If HasChanged Then Exit Function
    
    For i = 1 To UBound(Sections)
        HasChanged = (Sections(i).NewCode <> SecBackup(i).NewCode) Or (Sections(i).NewParams <> SecBackup(i).NewParams)
        If HasChanged Then Exit Function
    Next i

End Function

Private Sub cmbSection_Click()

    Dim i As Integer
    
    For i = 1 To UBound(Sections)
        If Sections(i).Name = cmbSection.Text Then
            txtCodeEditor(0).Text = Sections(i).NewCode
            txtCodeEditor(1).Text = Sections(i).code
            txtParams(0).Text = Sections(i).NewParams
            txtParams(1).Text = Sections(i).params
            txtDesc.Text = Sections(i).Description
            'imgIE.Visible = InStr(Sections(i).Support, "IE")
            'imgNS.Visible = InStr(Sections(i).Support, "NS")
            
            mnuCodeRemove.Enabled = Sections(i).IsNew
            mnuCodeEdit.Enabled = Sections(i).IsNew
            
            txtParams(0).Enabled = Sections(i).Name <> "GLOBAL" And Sections(i).Name <> "EVENT HANDLING"
            
            Exit For
        End If
    Next i
    
    On Error Resume Next
    
    txtCodeEditor(0).SetFocus

End Sub

Private Sub Form_Load()

    With pFont
        .Name = GetSetting(App.EXEName, "AddInEditorFont", "Name", "Terminal")
        .Size = GetSetting(App.EXEName, "AddInEditorFont", "Size", 9)
        .Bold = GetSetting(App.EXEName, "AddInEditorFont", "Bold", False)
        .Italic = GetSetting(App.EXEName, "AddInEditorFont", "Italic", False)
        .Underline = GetSetting(App.EXEName, "AddInEditorFont", "Underline", False)
    End With

    If Val(GetSetting(App.EXEName, "AddInEditorWinPos", "X")) = 0 Then
        CenterForm Me
    Else
        Top = GetSetting(App.EXEName, "AddInEditorWinPos", "X")
        Left = GetSetting(App.EXEName, "AddInEditorWinPos", "Y")
        Width = GetSetting(App.EXEName, "AddInEditorWinPos", "W")
        Height = GetSetting(App.EXEName, "AddInEditorWinPos", "H")
        WindowState = Val(GetSetting(App.EXEName, "AddInEditorWinPos", "State"))
    End If
    
    'SetupCharset Me
    
    OriginalAddIn = Project.AddIn
    
    UpdateCtrlsFont
    
    If Not IsDebug Then
        Set xMenu = New CMenu
        xMenu.Initialize Me
    End If

End Sub

Private Sub UpdateCtrlsFont()

    With pFont
        txtParams(0).FontName = .Name
        txtParams(0).FontSize = .Size
        txtParams(0).FontBold = .Bold
        txtParams(0).FontItalic = .Italic
        txtParams(0).FontUnderline = .Underline
    End With
    
    txtParams(1).Font = txtParams(0).Font
    txtCodeEditor(0).Font = txtParams(0).Font
    txtCodeEditor(1).Font = txtParams(0).Font

End Sub

Public Sub GetSections(AddIn As String)

    Dim sStr As String
    Dim FcnDone As Boolean
    Dim tmpName As String
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo InvalidFile
    
    If AddIn = "/default/" Then
        ReDim Sections(0)
        cmbSection.ComboItems.Clear
        mnuCode.Enabled = True
        mnuFileSave.Enabled = True
    End If
    i = 0

    ff = FreeFile
    If AddIn = "/default/" Or LenB(AddIn) = 0 Then
        Open AppPath + "rsc\code.dat" For Input As #ff
    Else
        Open AppPath + "AddIns\" + AddIn + ".ext" For Input As #ff
        Do Until sStr = "***"
            Line Input #ff, sStr
        Loop
    End If
    
        'Load Global Section
        i = i + 1
        If AddIn = "/default/" Then ReDim Preserve Sections(i)
        With Sections(i)
            .Name = "GLOBAL"
            .Support = "IE,NS"
            .Description = "Global definitions." + vbCrLf + "DO NOT REMOVE the comented lines (lines starting with %%)."
            Line Input #ff, sStr
            Do Until sStr = "%%TOOLBARSTYLE"
                If AddIn = "/default/" Then
                    .code = .code + Replace(sStr, vbTab, vbNullString) + vbCrLf
                Else
                    .NewCode = .NewCode + Replace(sStr, vbTab, vbNullString) + vbCrLf
                End If
                Line Input #ff, sStr
            Loop
            If AddIn = "/default/" Then
                .code = .code + Replace(sStr, vbTab, vbNullString)
                cmbSection.ComboItems.Add , , .Name, 1
            Else
                .NewCode = .NewCode + Replace(sStr, vbTab, vbNullString)
                If .code <> .NewCode Then cmbSection.ComboItems(i).Image = 2
            End If
        End With
    
        'Load Functions
        Do Until FcnDone
GetNextSection:
            Line Input #ff, sStr
            sStr = Replace$(sStr, vbTab, vbNullString)
            If Left$(sStr, 9) = "function " Then
                sStr = Mid$(sStr, 10)
                i = i + 1
                tmpName = GetFcnPart(sStr, 1, i)
                If AddIn = "/default/" Then
                    ReDim Preserve Sections(i)
                Else
                    Sections(i).ChangedFlag = csChangedFlag
                    j = i
                    Do While Sections(i).Name <> tmpName
                        If Sections(i).Name = "PrepareEvents" Then
                            i = j - 1
                            GetFcnPart sStr, 2
                            GetFcnPart sStr, 3
                            GetFcnPart sStr, 4
                            GetFcnPart sStr, 5
                            GoTo GetNextSection
                        End If
                        With Sections(i)
                            .NewParams = .params
                            .NewCode = .code
                        End With
                        i = i + 1
                    Loop
                End If
                With Sections(i)
                    .Name = tmpName
                    If AddIn = "/default/" Then
                        .params = GetFcnPart(sStr, 2)
                        .Support = GetFcnPart(sStr, 3)
                        .Description = GetFcnPart(sStr, 4)
                        .code = GetFcnPart(sStr, 5, i)
                        cmbSection.ComboItems.Add , , .Name, 1
                    Else
                        .NewParams = GetFcnPart(sStr, 2, i)
                        .Support = GetFcnPart(sStr, 3)
                        .Description = GetFcnPart(sStr, 4)
                        .NewCode = GetFcnPart(sStr, 5, i)
                        If .code <> .NewCode Then cmbSection.ComboItems(i).Image = 2
                    End If
                    If .Name = "PrepareEvents" Then
                        FcnDone = True
                        If .code = .NewCode Then cmbSection.ComboItems(i).Image = 1
                    End If
                End With
            End If
        Loop
        
        'Load Custom Functions
        If AddIn <> "/default/" And LenB(AddIn) <> 0 Then
            i = UBound(Sections)
            Do Until sStr = "%%TOOLBARCODE"
                Line Input #ff, sStr
                sStr = Replace$(sStr, vbTab, vbNullString)
                If Left$(sStr, 9) = "function " Then
                    sStr = Mid$(sStr, 10)
                    i = i + 1
                    ReDim Preserve Sections(i)
                    With Sections(i)
                        .Name = GetFcnPart(sStr, 1)
                        .params = GetFcnPart(sStr, 2, i): .NewParams = .params
                        .Support = GetFcnPart(sStr, 3)
                        .Description = GetFcnPart(sStr, 4)
                        .code = GetFcnPart(sStr, 5, i): .NewCode = .code
                        .IsNew = True
                        cmbSection.ComboItems.Add , , .Name, 3
                    End With
                End If
            Loop
            For j = 1 To UBound(Sections)
                If Sections(j).Name = "EVENT HANDLING" Then
                    i = j
                    Exit For
                End If
            Next j
        Else
            i = i + 1
        End If
    Close #ff
    
    If AddIn <> "/default/" Then
        SecBackup = Sections
        ParBackup = params
        LoadAddInParams AddIn
    End If
    
    DoEvents
    
    cmbSection.ComboItems(1).Selected = True
    cmbSection_Click
    
    Exit Sub
    
InvalidFile:
    MsgBox "Unable to read the sections." + vbCrLf + "The selected AddIn seems to be corrupted", vbCritical + vbOKOnly, "Error Loading Sections"
    On Error Resume Next
    cmbSection.ComboItems.Clear
    mnuCode.Enabled = False
    mnuFileSave.Enabled = False
    Close #ff

End Sub

Private Function GetFcnPart(ByVal sStr As String, Part As Integer, Optional i As Integer) As String

    Dim s As String
    Dim ss As String

    Select Case Part
        Case 1      'Name
            s = Left$(sStr, InStr(sStr, "(") - 1)
            If InStr(sStr, "/") > 0 Then
                csChangedFlag = Val(Mid(sStr, InStr(sStr, "/") - 1, 1))
                csIsOldAddIn = False
            Else
                csChangedFlag = False
                csIsOldAddIn = True
            End If
        Case 2      'Parameters
            s = Mid$(sStr, InStr(sStr, "(") + 1, InStr(sStr, ")") - InStr(sStr, "(") - 1)
            If Not csChangedFlag Then
                If Not csIsOldAddIn Then
                    If LenB(Sections(i).params) <> 0 Then s = Sections(i).params
                End If
            End If
        Case 3      'Support
            Line Input #ff, s
            s = Mid$(s, 5)
        Case 4      'Description
            Line Input #ff, s
            s = Mid$(s, 5) + vbCrLf
            Do Until EOF(ff)
                Line Input #ff, ss
                ss = Mid$(ss, 5)
                If LenB(ss) = 0 Then Exit Do
                s = s + ss + vbCrLf
            Loop
        Case 5      'Code
            Line Input #ff, s
            If Left$(s, 2) = "%%" Then
                s = s + vbCrLf
            Else
                s = Mid$(s, 3) + vbCrLf
            End If
            Do Until EOF(ff)
                Line Input #ff, ss
                If ss = vbTab + "}" Then Exit Do
                ss = Mid$(ss, 3)
                s = s + ss + vbCrLf
            Loop
            s = Left$(s, Len(s) - 2)
            
            If Not Sections(i).ChangedFlag Then
                If Not csIsOldAddIn Then
                    If LenB(Sections(i).code) <> 0 Then s = Sections(i).code
                End If
            End If
    End Select
    
    GetFcnPart = s
    
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    On Error GoTo ExitSub

    If HasChanged Then
        If Not Ask2Save Then
            Cancel = (Project.AddIn.Name <> "***")
            Exit Sub
        End If
    End If

    If WindowState = vbNormal Then
        SaveSetting App.EXEName, "AddInEditorWinPos", "X", Top
        SaveSetting App.EXEName, "AddInEditorWinPos", "Y", Left
        SaveSetting App.EXEName, "AddInEditorWinPos", "W", Width
        SaveSetting App.EXEName, "AddInEditorWinPos", "H", Height
    End If
    SaveSetting App.EXEName, "AddInEditorWinPos", "State", WindowState
    
    SaveSetting App.EXEName, "AddInEditorFont", "Name", pFont.Name
    SaveSetting App.EXEName, "AddInEditorFont", "Size", pFont.Size
    SaveSetting App.EXEName, "AddInEditorFont", "Bold", pFont.Bold
    SaveSetting App.EXEName, "AddInEditorFont", "Italic", pFont.Italic
    SaveSetting App.EXEName, "AddInEditorFont", "Underline", pFont.Underline
    
    Project.AddIn = OriginalAddIn
    
ExitSub:

End Sub

Private Function Ask2Save() As Boolean

    Dim Ans As Integer

    Ans = MsgBox("You have unsaved changes, would you like to save them now?", vbQuestion + vbYesNoCancel, "Save Changes")
    Select Case Ans
        Case vbYes
            Do Until LenB(Project.AddIn.Name) <> 0
                MsgBox "You must name this AddIn before saving it.", vbCritical + vbOKOnly, "AddIn Editor"
                frmAddInProp.Show vbModal
                If Project.AddIn.Name = "***" Then
                    Ask2Save = False
                    Exit Do
                End If
            Loop
            If Project.AddIn.Name <> "***" Then
                SaveAddIn
                Ask2Save = True
            End If
        Case vbNo
            Ask2Save = True
        Case vbCancel
            Ask2Save = False
    End Select

End Function

Private Sub Form_Resize()

    If WindowState = vbMinimized Then Exit Sub

    txtDesc.Top = Height - 700 - GetClientTop(Me.hwnd)
    txtDesc.Width = Width - 135

    txtCodeEditor(0).Height = Height - 1500 - GetClientTop(Me.hwnd)

    If mnuCodeViewOriginal.Checked Then
        txtCodeEditor(0).Width = Width / 2 - 105
    Else
        txtCodeEditor(0).Width = Width - 105
    End If
    
    txtParams(0).Width = txtCodeEditor(0).Width
    
    txtCodeEditor(1).Top = txtCodeEditor(0).Top
    txtCodeEditor(1).Height = txtCodeEditor(0).Height
    txtCodeEditor(1).Left = txtCodeEditor(0).Width + 75
    txtCodeEditor(1).Width = txtCodeEditor(0).Width
    
    txtParams(1).Top = txtParams(0).Top
    txtParams(1).Left = txtCodeEditor(1).Left
    txtParams(1).Width = txtCodeEditor(1).Width
    txtParams(1).Height = txtParams(0).Height

End Sub

Private Sub mnuCodeAdd_Click()

    Dim nf As Integer
    
    nf = UBound(Sections)
    frmAddFunction.Show vbModal
    If nf < UBound(Sections) Then
        PropChanged = True
        With Sections(UBound(Sections))
            cmbSection.ComboItems.Add , , .Name, 3
        End With
        cmbSection.ComboItems.item(UBound(Sections)).Selected = True
        DoEvents
        cmbSection_Click
    End If

End Sub

Private Sub mnuCodeEdit_Click()

    Dim i As Integer

    With frmAddFunction
        For i = 1 To UBound(Sections)
            If Sections(i).Name = cmbSection.Text Then
                .txtName.Text = Sections(i).Name
                .txtName.tag = Sections(i).Name
                .txtDesc.Text = Sections(i).Description
                '.chkIE.Value = Abs(InStr(Sections(i).Support, "IE") > 0)
                '.chkNS.Value = Abs(InStr(Sections(i).Support, "NS") > 0)
                .cmdOK.caption = "&Update"
                .Show vbModal
                cmbSection.SelectedItem.Text = Sections(i).Name
                Exit For
            End If
        Next i
    End With

    PropChanged = True

End Sub

Private Sub mnuCodeRemove_Click()

    Dim i As Integer
    Dim j As Integer
    
    If MsgBox("Are you sure you want to remove the function " + cmbSection.Text + "?", vbQuestion + vbYesNo, "Confirm Remove") = vbYes Then
        For i = 1 To UBound(Sections)
            If Sections(i).Name = cmbSection.Text Then
                PropChanged = True
                For j = i To UBound(Sections) - 1
                    Sections(j) = Sections(j + 1)
                Next j
                ReDim Preserve Sections(UBound(Sections) - 1)
                cmbSection.ComboItems.Remove cmbSection.SelectedItem.Index
                cmbSection.ComboItems(1).Selected = True
                cmbSection_Click
                Exit For
            End If
        Next i
    End If
    
End Sub

Private Sub mnuCodeRestore_Click()

    With cmbSection.SelectedItem
        Sections(.Index).NewCode = Sections(.Index).code
        Sections(.Index).NewParams = Sections(.Index).params
        txtCodeEditor(0).Text = Sections(.Index).NewCode
        txtParams(0).Text = Sections(.Index).NewParams
    End With
    
    CheckChanges

End Sub

Private Sub mnuCodeViewOriginal_Click()

    mnuCodeViewOriginal.Checked = Not mnuCodeViewOriginal.Checked
    Form_Resize

End Sub

Private Sub mnuFileExit_Click()

    Unload Me

End Sub

Private Sub mnuFileNew_Click()

    If HasChanged Then
        If Not Ask2Save Then Exit Sub
    End If

    PropChanged = True
    GetSections "/default/"
    GetSections vbNullString
    Project.AddIn.Name = vbNullString
    Project.AddIn.Description = vbNullString
    frmAddInProp.Show vbModal
    UpdateTitleBar

End Sub

Private Sub mnuFileOpen_Click()

    If HasChanged Then
        If Not Ask2Save Then Exit Sub
    End If

    frmOpenAddIn.Show vbModal

End Sub

Private Sub mnuFileProperties_Click()

    Dim oName As String
    
    oName = Project.AddIn.Name

    frmAddInProp.Show vbModal
    
    If Project.AddIn.Name = "***" Then
        Project.AddIn.Name = oName
    Else
        PropChanged = True
        UpdateTitleBar
    End If

End Sub

Public Sub UpdateTitleBar()

    With Project.AddIn
        If LenB(.Name) <> 0 Then
            Me.caption = "AddIn Editor 1.5 - [" + Project.AddIn.Name + "]"
        Else
            Me.caption = "AddIn Editor 1.5 - [Untitled]"
        End If
    End With

End Sub

Private Sub mnuFileSave_Click()

    If LenB(Project.AddIn.Name) = 0 Then
        MsgBox "You must name this AddIn. Select Properties from the File menu to set a name for this AddIn", vbCritical + vbOKOnly, "AddIn Editor"
    Else
        SaveAddIn
    End If
    
End Sub

Private Sub SaveAddIn()

    Dim i As Integer
    Dim tmpStr As String
    
    On Error GoTo Unable2Save

    ff = FreeFile
    Open AppPath + "AddIns\" + Project.AddIn.Name + ".ext" For Output As #ff
        Print #ff, Project.AddIn.Description
        Print #ff, "***"
        
        'Save Global section
        tmpStr = vbTab + Replace(Sections(1).NewCode, vbCrLf, vbCrLf + vbTab)
        tmpStr = Replace$(tmpStr, vbTab + "%%", "%%")
        Print #ff, tmpStr + vbCrLf
        
        'Save Functions
        For i = 2 To UBound(Sections)
            Print #ff, vbTab + "function " + Sections(i).Name + "(" + Sections(i).NewParams + ")" + IIf(Sections(i).code = Sections(i).NewCode, "0", "1") + "/ {"
            Print #ff, String$(2, vbTab) + "//" + Replace(Sections(i).Support, vbCrLf, vbCrLf + String$(2, vbTab) + "//")
            Print #ff, String$(2, vbTab) + "//" + Replace(Sections(i).Description, vbCrLf, vbCrLf + String$(2, vbTab) + "//")
            Print #ff, String$(2, vbTab) + Replace(Sections(i).NewCode, vbCrLf, vbCrLf + String$(2, vbTab))
            Print #ff, vbTab + "}" + vbCrLf
        Next i
        
        'Save Event Handling
        Print #ff, "%%TOOLBARCODE" + vbCrLf
        Print #ff, "%%KBDNAVSUP" + vbCrLf
        Print #ff, "%%BROWSERCODE"
    Close #ff
    
    PropChanged = False
    
    Exit Sub
    
Unable2Save:
InvalidFile:
    MsgBox "Unable to save the sections", vbCritical + vbOKOnly, "AddIn Editor"
    On Error Resume Next
    Close #ff

End Sub

Private Sub mnuHelpAbout_Click()

    MsgBox Me.caption + vbCrLf + vbCrLf + "Resources Path:" + vbCrLf + AppPath + "rsc\"

End Sub

Private Sub mnuHelpContents_Click()

    showHelp "aie.htm"

End Sub

Private Sub mnuOpFont_Click()

    On Error GoTo ExitSub
    
    With cDlg
        .DialogTitle = "Select Font"
        .CancelError = True
        .Flags = cdlCFBoth
        
        .FontName = pFont.Name
        .FontSize = pFont.Size
        .FontBold = pFont.Bold
        .FontItalic = pFont.Italic
        .FontUnderline = pFont.Underline
        
        .ShowFont
        
        pFont.Name = .FontName
        pFont.Size = .FontSize
        pFont.Bold = .FontBold
        pFont.Italic = .FontItalic
        pFont.Underline = .FontUnderline
    End With
    
    UpdateCtrlsFont
    
ExitSub:
    Exit Sub

End Sub

Private Sub mnuParamAdd_Click()

    With frmParamMan
        .caption = "Add Parameter"
        .picContainer.Left = 45
        .cmdOK.caption = "&Add"
        .cmdOK.Left = 1350
        .cmdCancel.Left = 2445
        .cmdDelete.Visible = False
        .Width = 3500
        .Show vbModal
    End With
    
    SaveAddInParams Project.AddIn.Name

End Sub

Private Sub mnuParamEdit_Click()

    With frmParamMan
        .caption = "Edit Parameters"
        .picContainer.Left = 2250
        .cmdOK.caption = "&Update"
        .cmdOK.Left = 3480
        .cmdCancel.Left = 4575
        .Width = 5715
        
        .txtName.Enabled = False
        .txtDesc.Enabled = False
        .txtDefault.Enabled = False
        .chkRequired.Enabled = False
        .cmdDelete.Enabled = False

        .Show vbModal
    End With
    
    SaveAddInParams Project.AddIn.Name

End Sub

Private Sub mnuParamInsert_Click()

    frmParamSel.Show vbModal
    CheckChanges

End Sub

Private Sub tmrInit_Timer()

    tmrInit.Enabled = False
    GetSections "/default/"
    If LenB(Project.AddIn.Name) <> 0 Then
        If Not FileExists(AppPath + "AddIns\" + Project.AddIn.Name + ".ext") Then
            MsgBox "The AddIn " + Project.AddIn.Name + " could not be found." + vbCrLf + "AddIn Editor will start with a blank project", vbInformation + vbOKOnly, "Error Opening AddIn"
            Project.AddIn.Name = vbNullString
        End If
    End If
    GetSections Project.AddIn.Name
    UpdateTitleBar

End Sub

Private Sub txtCodeEditor_GotFocus(Index As Integer)

    Dim c As Control
    
    On Error Resume Next
    
    For Each c In Me.Controls
        c.TabStop = False
    Next c

End Sub

Private Sub txtCodeEditor_LostFocus(Index As Integer)

    Dim c As Control
    
    On Error Resume Next
    
    For Each c In Me.Controls
        If (TypeOf c Is Button Or TypeOf c Is TextBox Or TypeOf c Is TextBox) And c.Enabled And c.Visible Then c.TabStop = True
    Next c

End Sub

Private Sub txtCodeEditor_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyA And Shift = vbCtrlMask Then
        txtCodeEditor(0).SelStart = 0
        txtCodeEditor(0).SelLength = Len(txtCodeEditor(0).Text)
    Else
        CheckChanges
    End If

End Sub

Private Sub CheckChanges()

    Dim i As Integer
    
    For i = 1 To UBound(Sections)
        If Sections(i).Name = cmbSection.SelectedItem.Text Then
            Exit For
        End If
    Next i

    Sections(i).NewCode = txtCodeEditor(0).Text
    Sections(i).NewParams = txtParams(0).Text
    cmbSection.SelectedItem.Image = (1 + 2 * Abs(Sections(i).IsNew)) + _
                Abs((Sections(i).code <> Sections(i).NewCode) Or _
                (Sections(i).params <> Sections(i).NewParams))

End Sub

Private Sub txtParams_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)

    CheckChanges

End Sub

#End If

