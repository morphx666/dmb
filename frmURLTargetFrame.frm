VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmURLTargetFrame 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Target Frame"
   ClientHeight    =   4455
   ClientLeft      =   6630
   ClientTop       =   7200
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
   Icon            =   "frmURLTargetFrame.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   5025
   ShowInTaskbar   =   0   'False
   Begin SmartButtonProject.SmartButton cmdReload 
      Height          =   315
      Left            =   4590
      TabIndex        =   3
      Top             =   330
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmURLTargetFrame.frx":014A
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
   Begin SmartButtonProject.SmartButton cmdBrowseFramesDoc 
      Height          =   315
      Left            =   4215
      TabIndex        =   2
      Top             =   330
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmURLTargetFrame.frx":04E4
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
   Begin VB.TextBox txtFramesDocument 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   75
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   345
      Width           =   4065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3015
      TabIndex        =   10
      Top             =   4005
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4050
      TabIndex        =   11
      Top             =   4005
      Width           =   900
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3855
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   53
   End
   Begin VB.TextBox txtFrame 
      Height          =   285
      Left            =   75
      TabIndex        =   8
      Top             =   3405
      Width           =   4875
   End
   Begin VB.ListBox lstFrames 
      Height          =   1815
      Left            =   75
      TabIndex        =   6
      Top             =   1245
      Width           =   4875
   End
   Begin VB.Label lblFramesInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "..."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   75
      TabIndex        =   4
      Top             =   675
      Width           =   90
   End
   Begin VB.Label lblFramesDoc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Frames Document"
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   105
      Width           =   1290
   End
   Begin VB.Label lblTFrame 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Target Frame"
      Height          =   195
      Left            =   75
      TabIndex        =   7
      Top             =   3165
      Width           =   975
   End
   Begin VB.Label lblAFrames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Available Frames"
      Height          =   195
      Left            =   75
      TabIndex        =   5
      Top             =   1005
      Width           =   1215
   End
End
Attribute VB_Name = "frmURLTargetFrame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBrowseFramesDoc_Click()

    On Error GoTo ExitSub
    
    With frmMain.cDlg
        .DialogTitle = GetLocalizedStr(385)
        .InitDir = GetRealLocal.RootWeb
        .filter = SupportedHTMLDocs
        .CancelError = True
        .Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNPathMustExist
        .ShowOpen
        txtFramesDocument.Text = .FileName
        cmdReload_Click
    End With
    
ExitSub:
    Exit Sub

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    If LenB(txtFrame.Text) = 0 Then
        txtFrame.Text = "_self"
    Else
        Select Case LCase(txtFrame.Text)
            Case "self": txtFrame.Text = "_self"
            Case "top": txtFrame.Text = "_top"
            Case "blank": txtFrame.Text = "_blank"
            Case "parent": txtFrame.Text = "_parent"
        End Select
    End If

    frmMain.cmbTargetFrame.Text = txtFrame.Text
    Unload Me

End Sub

Private Sub cmdReload_Click()

    InitDlg

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/url_targetframe.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    SetupCharset Me
    LocalizeUI

    txtFramesDocument.Text = Project.UserConfigs(Project.DefaultConfig).Frames.FramesFile

    InitDlg

End Sub

Private Sub InitDlg()

    Dim i As Integer
    Dim nf As Integer
    Dim fn As String
    Dim oFramesFile As String
    
    On Error Resume Next
    
    oFramesFile = FramesInfo.FileName
    FramesInfo.FileName = txtFramesDocument.Text
    GetFramesInfo
    FramesInfo.FileName = oFramesFile
    
    With lstFrames
        .Clear
        AddFrame2List "_self"
        AddFrame2List "_top"
        AddFrame2List "_blank"
        AddFrame2List "_parent"
    End With
    
    If FramesInfo.IsValid Then
        nf = UBound(FramesInfo.Frames)
        For i = 1 To nf
            AddFrame2List FramesInfo.Frames(i).Name
        Next i
        lblFramesInfo.caption = GetLocalizedStr(888) + " " & nf & " " + IIf(nf > 1, GetLocalizedStr(890), GetLocalizedStr(889))
    Else
        If FileExists(txtFramesDocument.Text) Then
            lblFramesInfo.caption = GetLocalizedStr(886)
        Else
            lblFramesInfo.caption = GetLocalizedStr(887)
        End If
    End If
    
    For i = 1 To UBound(MenuGrps)
        AddFrame2List MenuGrps(i).Actions.onclick.TargetFrame
        AddFrame2List MenuGrps(i).Actions.onmouseover.TargetFrame
        AddFrame2List MenuGrps(i).Actions.OnDoubleClick.TargetFrame
    Next i
    
    For i = 1 To UBound(MenuCmds)
        AddFrame2List MenuCmds(i).Actions.onclick.TargetFrame
        AddFrame2List MenuCmds(i).Actions.onmouseover.TargetFrame
        AddFrame2List MenuCmds(i).Actions.OnDoubleClick.TargetFrame
    Next i
    
    If IsCommand(frmMain.tvMenus.SelectedItem.key) Then
        Select Case frmMain.tsCmdType.SelectedItem.key
            Case "tsOver"
                fn = MenuCmds(GetID).Actions.onmouseover.TargetFrame
            Case "tsClick"
                fn = MenuCmds(GetID).Actions.onclick.TargetFrame
            Case "tsDoubleClick"
                fn = MenuCmds(GetID).Actions.OnDoubleClick.TargetFrame
        End Select
    Else
        Select Case frmMain.tsCmdType.SelectedItem.key
            Case "tsOver"
                fn = MenuGrps(GetID).Actions.onmouseover.TargetFrame
            Case "tsClick"
                fn = MenuGrps(GetID).Actions.onclick.TargetFrame
            Case "tsDoubleClick"
                fn = MenuGrps(GetID).Actions.OnDoubleClick.TargetFrame
        End Select
    End If
    fn = SetNiceFrameName(fn)
    
    For i = 0 To lstFrames.ListCount - 1
        If fn = lstFrames.List(i) Then
            lstFrames.ListIndex = i
            Exit For
        End If
    Next i
    
    caption = GetLocalizedStr(235)
    fn = Project.UserConfigs(Project.DefaultConfig).Frames.FramesFile
    If LenB(fn) <> 0 Then
        If FileExists(fn) Then
            caption = caption + " - " + GetFileName(fn)
        End If
    End If
    
End Sub

Private Sub AddFrame2List(ByVal fn As String)

    If LenB(fn) <> 0 Then
        fn = SetNiceFrameName(fn)
        If Not Exists(fn) Then
            lstFrames.AddItem fn
        End If
    End If
    
End Sub

Private Function Exists(fn As String) As Boolean

    Dim i As Integer
    
    For i = 0 To lstFrames.ListCount - 1
        If lstFrames.List(i) = fn Then
            Exists = True
            Exit Function
        End If
    Next i

End Function

Private Sub lstFrames_Click()

    txtFrame.Text = GetFrameNameFromList(lstFrames.ListIndex)

End Sub

Private Function SetNiceFrameName(ByVal fn As String) As String

    Select Case fn
        Case "_self", "self": fn = "Same Frame (_self)"
        Case "_top", "top": fn = "Whole Page (_top)"
        Case "_blank", "blank": fn = "New Window (_blank)"
        Case "_parent", "parent": fn = "Parent Frame (_parent)"
    End Select
    
    SetNiceFrameName = fn

End Function

Private Function GetFrameNameFromList(idx As Integer) As String

    Dim fn As String

    Select Case idx
        Case 0: fn = "_self"
        Case 1: fn = "_top"
        Case 2: fn = "_blank"
        Case 3: fn = "_parent"
        Case Is > 3
            fn = lstFrames.List(idx)
    End Select
    
    GetFrameNameFromList = fn

End Function

Private Sub lstFrames_DblClick()

    If lstFrames.ListIndex >= 0 Then cmdOK_Click

End Sub

Private Sub txtFrame_GotFocus()

    SelAll txtFrame

End Sub

Private Sub txtFramesDocument_Change()

    cmdReload.Enabled = FileExists(txtFramesDocument.Text)

End Sub

Private Sub LocalizeUI()

    lblFramesDoc.caption = GetLocalizedStr(884)
    lblAFrames.caption = GetLocalizedStr(885)
    lblTFrame.caption = GetLocalizedStr(235)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
