VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmStyleGeneral 
   Caption         =   "General"
   ClientHeight    =   4845
   ClientLeft      =   4260
   ClientTop       =   5715
   ClientWidth     =   7500
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4845
   ScaleWidth      =   7500
   Begin VB.TextBox txtURL 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   315
      Left            =   1005
      TabIndex        =   1
      ToolTipText     =   "Type the URL address of the link you want the browser to follow when this item is triggered"
      Top             =   630
      WhatsThisHelpID =   20070
      Width           =   3135
   End
   Begin VB.CheckBox chkEnabled 
      Alignment       =   1  'Right Justify
      Caption         =   "Enabled"
      Height          =   225
      Left            =   45
      TabIndex        =   6
      ToolTipText     =   "Check this option to enable events for this command"
      Top             =   1545
      Width           =   1185
   End
   Begin VB.TextBox txtCaption 
      Height          =   495
      Left            =   1005
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   60
      Width           =   3135
   End
   Begin VB.TextBox txtStatus 
      Height          =   315
      Left            =   1005
      TabIndex        =   5
      Top             =   1020
      Width           =   3135
   End
   Begin SmartButtonProject.SmartButton cmdBrowse 
      Height          =   315
      Left            =   4170
      TabIndex        =   2
      Top             =   630
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmGeneral.frx":0000
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
   Begin SmartButtonProject.SmartButton cmdTargetFrame 
      Height          =   315
      Left            =   4950
      TabIndex        =   4
      Top             =   630
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmGeneral.frx":015A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SmartButtonProject.SmartButton cmdBookmark 
      Height          =   315
      Left            =   4560
      TabIndex        =   3
      Top             =   630
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmGeneral.frx":02B4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblActionName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Link"
      Height          =   195
      Left            =   45
      TabIndex        =   9
      Top             =   690
      Width           =   270
   End
   Begin VB.Label lblStatus 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Status Text"
      Height          =   195
      Left            =   45
      TabIndex        =   8
      Top             =   1080
      Width           =   840
   End
   Begin VB.Label lblCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      Height          =   195
      Left            =   45
      TabIndex        =   7
      Top             =   120
      Width           =   555
   End
End
Attribute VB_Name = "frmStyleGeneral"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsUpdating As Boolean

Private Sub chkEnabled_Click()

    If IsUpdating Then Exit Sub
    UpdateMenuItem

End Sub

Private Sub UpdateMenuItem()

    frmMain.UpdateItemData GetLocalizedStr(189) + cSep + GetLocalizedStr(232), False

End Sub

Private Sub Form_Load()

    LocalizeUI

End Sub

Friend Sub UpdateUI(c As MenuCmd, IsC As Boolean, IsG As Boolean)

    IsUpdating = True
    
    lblCaption.Enabled = (IsG Or IsC) And Not c.disabled
    txtCaption.Enabled = (IsG Or IsC) And Not c.disabled
    lblStatus.Enabled = (IsG Or IsC) And Not c.disabled
    txtStatus.Enabled = (IsG Or IsC) And Not c.disabled
    chkEnabled.Enabled = (IsG Or IsC)
    lblActionName.Enabled = (IsG Or IsC) And Not c.disabled
    txtURL.Enabled = (IsG Or IsC) And Not c.disabled
    cmdBookmark.Enabled = (IsG Or IsC) And Not c.disabled
    cmdBrowse.Enabled = (IsG Or IsC) And Not c.disabled
    cmdTargetFrame.Enabled = (IsG Or IsC) And Not c.disabled
    
    txtCaption.Text = Replace(c.Caption, "<br>", vbCrLf, , , vbTextCompare)
    txtStatus.Text = c.WinStatus
    txtURL.Text = c.Actions.onclick.URL
    chkEnabled.Value = IIf(c.disabled, vbUnchecked, vbChecked)
    
    IsUpdating = False
    
End Sub

Private Sub cmdBookmark_Click()

    frmURLBookmark.Show vbModal
    SetBookmarkState

End Sub

Friend Sub cmdBrowse_Click()

    Dim ActionName As String
    Dim FileName As String
    Dim sStr As String
    Dim ItemName As String
    
    If IsGroup(frmMain.tvMenus.SelectedItem.key) Then
        ItemName = NiceGrpCaption(GetID)
    Else
        ItemName = NiceCmdCaption(GetID)
    End If
    ItemName = "'" + ItemName + "'"

    If LenB(InitDir) = 0 Or Left(InitDir, Len(GetRealLocal.RootWeb)) <> GetRealLocal.RootWeb Then
        InitDir = GetRealLocal.RootWeb
    End If
    With frmMain.cDlg
        ActionName = ItemName + " " + GetLocalizedStr(236)
        sStr = GetFilePath(Replace(txtURL.Text, Project.UserConfigs(Project.DefaultConfig).RootWeb, GetRealLocal.RootWeb))
        If FolderExists(sStr) And LenB(sStr) <> 0 Then
            InitDir = sStr
            FileName = GetFileName(txtURL.Text)
        End If
        
        .DialogTitle = GetLocalizedStr(239) + " " + ActionName
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames '+ cdlOFNNoReadOnlyReturn
        .filter = SupportedHTMLDocs
        .InitDir = InitDir
        .FileName = FileName
        Err.Clear
        On Error Resume Next
        .ShowOpen
        If .FileName = FileName And LenB(FileName) <> 0 Then .FileName = InitDir + FileName
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then Exit Sub
        txtURL.Text = ConvertPath(.FileName)
        InitDir = GetFilePath(.FileName)
    End With

End Sub

Friend Sub Form_Resize()

    On Error Resume Next
    
    Dim m As Single
    
    m = lblCaption.Left * 2 + 75 * 15

    txtCaption.Width = Width - m
    txtStatus.Width = txtCaption.Width
    
    cmdTargetFrame.Left = txtCaption.Left + txtCaption.Width - cmdTargetFrame.Width
    cmdBookmark.Left = cmdTargetFrame.Left - cmdTargetFrame.Width - 15 * 2
    cmdBrowse.Left = cmdBookmark.Left - cmdTargetFrame.Width - 15 * 2
    txtURL.Width = cmdBrowse.Left - 15 * 2 - txtURL.Left

End Sub

Private Sub txtURL_Change()

    If IsUpdating Then Exit Sub
    DontRefreshMap = True
    SetBookmarkState
    frmMain.UpdateItemData GetLocalizedStr(189) + " " + cSep + " " + GetLocalizedStr(294)

    Dim si As Node
    Dim n As Node
    Dim id As Integer

    If InMapMode Then
        Set si = frmMain.tvMenus.SelectedItem
        Set n = frmMain.tvMapView.SelectedItem
        id = GetID(si)
        If IsGroup(si.key) Then
            If MemberOf(id) = 0 Then
                n.ForeColor = frmMain.GetItemColorByURL(MenuGrps(id).Actions, Preferences.GroupStyle.Color)
            Else
                n.ForeColor = frmMain.GetItemColorByURL(MenuGrps(id).Actions, Preferences.ToolbarItemStyle.Color)
            End If
        Else
            n.ForeColor = frmMain.GetItemColorByURL(MenuCmds(id).Actions, Preferences.CommandStyle.Color)
        End If
    End If
    
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)

    If Len(frmStyleGeneral.txtURL.Text) > 11 Then
        If Not (IsExternalLink(frmStyleGeneral.txtURL.Text) Or UsesProtocol(frmStyleGeneral.txtURL.Text)) Then
            With TipsSys
                .CanDisable = True
                .TipTitle = GetLocalizedStr(877)
                .Tip = GetLocalizedStr(878)
                .Show
            End With
        End If
    End If

End Sub

Private Sub txtStatus_Change()

    If IsUpdating Then Exit Sub
    DontRefreshMap = True
    frmMain.UpdateItemData GetLocalizedStr(234) + " " + cSep + " " + GetLocalizedStr(114)
    
End Sub

Private Sub txtCaption_Change()

    If IsUpdating Then Exit Sub
    
    DontRefreshMap = True
    
    frmMain.UpdateItemData GetLocalizedStr(234) + " " + cSep + " " + GetLocalizedStr(103), True, False
    
End Sub

Private Sub cmdTargetFrame_Click()

    frmURLTargetFrame.Show vbModal
    SetBookmarkState

End Sub

Private Sub SetBookmarkState()

    Dim fn As String
    Dim miActions As ActionEvents

    frmStyleGeneral.cmdBookmark.Enabled = LenB(frmStyleGeneral.txtURL.Text) <> 0 And Not UsesProtocol(frmStyleGeneral.txtURL.Text) And frmStyleGeneral.txtURL.Enabled

    frmStyleGeneral.cmdBookmark.ToolTipText = "Bookmark"
    If frmStyleGeneral.cmdBookmark.Enabled And frmStyleGeneral.cmdBookmark.Visible Then
        fn = frmStyleGeneral.txtURL.Text
        If InStr(fn, "#") Then
            frmStyleGeneral.cmdBookmark.ToolTipText = frmStyleGeneral.cmdBookmark.ToolTipText + " (" + Mid(fn, InStrRev(fn, "#") + 1) + ")"
        End If
    End If
    
    fn = ""
    frmStyleGeneral.cmdTargetFrame.ToolTipText = GetLocalizedStr(851)
    If IsGroup(frmMain.tvMenus.SelectedItem.key) Then
        miActions = MenuGrps(GetID).Actions
    Else
        miActions = MenuCmds(GetID).Actions
    End If
    With miActions
        If .onclick.Type = atcURL Then fn = .onclick.TargetFrame
    End With
    If LenB(fn) <> 0 Then frmStyleGeneral.cmdTargetFrame.ToolTipText = frmStyleGeneral.cmdTargetFrame.ToolTipText + " (" + fn + ")"
    
End Sub

Private Sub LocalizeUI()

    lblCaption.Caption = GetLocalizedStr(103)
    chkEnabled.Caption = GetLocalizedStr(104)
    
    lblStatus.Caption = GetLocalizedStr(114)

End Sub
