VERSION 5.00
Begin VB.Form frmPPImportOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import Porject Options"
   ClientHeight    =   4680
   ClientLeft      =   6930
   ClientTop       =   3780
   ClientWidth     =   5460
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5460
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4485
      TabIndex        =   11
      Top             =   4230
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3360
      TabIndex        =   10
      Top             =   4230
      Width           =   900
   End
   Begin VB.Frame Frame2 
      Caption         =   "Style"
      Height          =   1080
      Left            =   60
      TabIndex        =   3
      Top             =   3000
      Width           =   4560
      Begin VB.ComboBox cmbCmds 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   510
         Width           =   1815
      End
      Begin VB.ComboBox cmbGrps 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   285
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   510
         Width           =   1815
      End
      Begin VB.Label cmbCommands 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Command"
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
         Left            =   2400
         TabIndex        =   6
         Top             =   285
         Width           =   705
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
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
         Left            =   285
         TabIndex        =   4
         Top             =   285
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Project Properties"
      Height          =   2445
      Left            =   60
      TabIndex        =   1
      Top             =   465
      Width           =   4560
      Begin VB.CheckBox chkToolbar 
         Caption         =   "Toolbar Settings"
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
         Left            =   255
         TabIndex        =   13
         Top             =   2055
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin VB.CheckBox chkConfigs 
         Caption         =   "Configurations"
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
         Left            =   255
         TabIndex        =   12
         Top             =   270
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin VB.CheckBox chkAdvanced 
         Caption         =   "Advanced Settings"
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
         Left            =   255
         TabIndex        =   9
         Top             =   1785
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin VB.CheckBox chkFTP 
         Caption         =   "FTP Settings"
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
         Left            =   255
         TabIndex        =   8
         Top             =   1515
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin VB.ListBox lstConfigs 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   510
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   495
         Width           =   3660
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Select which options you want to import from the project"
      Height          =   210
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   4800
   End
End
Attribute VB_Name = "frmPPImportOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SrcProject As ProjectDef
Dim IsUpdating As Boolean

Private Sub chkConfigs_Click()

    lstConfigs.Enabled = (chkConfigs.Value = vbChecked)

End Sub

Private Sub cmdCancel_Click()

    Unload Me

End Sub

Private Sub cmdOK_Click()

    Dim i As Integer
    Dim ff As Integer
    Dim sCode As String
    Dim fCode As String
    Dim sLine As String
    Dim sGrp As MenuGrp
    Dim sCmd As MenuCmd
    Dim DoneStatus As Integer
    
    On Error GoTo ReportError

    If chkConfigs.Value = vbChecked Then
        For i = 0 To lstConfigs.ListCount - 1
            If i > UBound(Project.UserConfigs) Then
                ReDim Preserve Project.UserConfigs(UBound(Project.UserConfigs) + 1)
                Project.UserConfigs(UBound(Project.UserConfigs)).Name = SrcProject.UserConfigs(i).Name
            End If
            Project.UserConfigs(GetConfigID(SrcProject.UserConfigs(i).Name)) = SrcProject.UserConfigs(i)
        Next i
    End If
    
    If chkFTP.Value = vbChecked Then
        Project.FTP = SrcProject.FTP
    End If
    
    If chkAdvanced.Value = vbChecked Then
        Project.CodeOptimization = SrcProject.CodeOptimization
        Project.AddIn = SrcProject.AddIn
        Project.JSFileName = SrcProject.JSFileName
    End If
    
    If chkToolbar.Value = vbChecked Then
        Project.ToolBar = SrcProject.ToolBar
        ReDim Project.ToolBar.Groups(0)
    End If
    
    sCode = LoadFile(ImportProjectFileName)
    If InStr(sCode, "[RSC]") Then
        sCode = Split(sCode, "[RSC]")(1)
        sCode = LoadFile(Project.FileName)
        If InStr(fCode, "[RSC]") = 0 Then
            If Right(fCode, 1) <> vbCrLf Then fCode = fCode + vbCrLf
            fCode = fCode + "[RSC]"
        End If
        fCode = fCode + sCode
        ff = FreeFile
        Open Project.FileName For Output As #ff
            Print #ff, fCode
        Close #ff
    End If
    
    If cmbGrps.ListIndex = 0 Then
        AddMenuGroup "[Import]"
    End If
    
    ff = FreeFile
    Open ImportProjectFileName For Input As #ff
        Do Until EOF(ff) Or DoneStatus = 2
            Input #ff, sLine
            If Left(sLine, 3) = "[G]" And cmbGrps.ListIndex > 0 Then
                If Mid(GetParam(sLine, 1), 4) = cmbGrps.Text Then
                    AddMenuGroup Mid(sLine, 5)
                    For i = 1 To UBound(MenuGrps) - 1
                        MenuGrps(i).Alignment = MenuGrps(UBound(MenuGrps)).Alignment
                        MenuGrps(i).BackImage = MenuGrps(UBound(MenuGrps)).BackImage
                        MenuGrps(i).bcolor = MenuGrps(UBound(MenuGrps)).bcolor
                        MenuGrps(i).CaptionAlignment = MenuGrps(UBound(MenuGrps)).CaptionAlignment
                        MenuGrps(i).CmdsFXhColor = MenuGrps(UBound(MenuGrps)).CmdsFXhColor
                        MenuGrps(i).CmdsFXnColor = MenuGrps(UBound(MenuGrps)).CmdsFXnColor
                        MenuGrps(i).CmdsFXNormal = MenuGrps(UBound(MenuGrps)).CmdsFXNormal
                        MenuGrps(i).CmdsFXOver = MenuGrps(UBound(MenuGrps)).CmdsFXOver
                        MenuGrps(i).CmdsFXSize = MenuGrps(UBound(MenuGrps)).CmdsFXSize
                        MenuGrps(i).CmdsMarginX = MenuGrps(UBound(MenuGrps)).CmdsMarginX
                        MenuGrps(i).CmdsMarginY = MenuGrps(UBound(MenuGrps)).CmdsMarginY
                        MenuGrps(i).ContentsMarginH = MenuGrps(UBound(MenuGrps)).ContentsMarginH
                        MenuGrps(i).ContentsMarginV = MenuGrps(UBound(MenuGrps)).ContentsMarginV
                        MenuGrps(i).Corners = MenuGrps(UBound(MenuGrps)).Corners
                        MenuGrps(i).DefHoverFont = MenuGrps(UBound(MenuGrps)).DefHoverFont
                        MenuGrps(i).DefNormalFont = MenuGrps(UBound(MenuGrps)).DefNormalFont
                        MenuGrps(i).DropShadow = MenuGrps(UBound(MenuGrps)).DropShadow
                        MenuGrps(i).frameBorder = MenuGrps(UBound(MenuGrps)).frameBorder
                        MenuGrps(i).hBackColor = MenuGrps(UBound(MenuGrps)).hBackColor
                        MenuGrps(i).HSImage = MenuGrps(UBound(MenuGrps)).HSImage
                        MenuGrps(i).hTextColor = MenuGrps(UBound(MenuGrps)).hTextColor
                        MenuGrps(i).iCursor = MenuGrps(UBound(MenuGrps)).iCursor
                        MenuGrps(i).Image = MenuGrps(UBound(MenuGrps)).Image
                        MenuGrps(i).Leading = MenuGrps(UBound(MenuGrps)).Leading
                        MenuGrps(i).LeftImage = MenuGrps(UBound(MenuGrps)).LeftImage
                        MenuGrps(i).nBackColor = MenuGrps(UBound(MenuGrps)).nBackColor
                        MenuGrps(i).nTextColor = MenuGrps(UBound(MenuGrps)).nTextColor
                        MenuGrps(i).RightImage = MenuGrps(UBound(MenuGrps)).RightImage
                        MenuGrps(i).Sound = MenuGrps(UBound(MenuGrps)).Sound
                        MenuGrps(i).Transparency = MenuGrps(UBound(MenuGrps)).Transparency
                    Next i
                    DoneStatus = DoneStatus + 1
                End If
            End If
            If Left(sLine, 3) = "[C]" And cmbCmds.ListIndex > 0 Then
                If Mid(GetParam(sLine, 1), 6) = cmbCmds.Text Then
                    AddMenuCommand Mid(sLine, 6)
                    For i = 1 To UBound(MenuCmds)
                        MenuCmds(i).Alignment = MenuCmds(UBound(MenuCmds)).Alignment
                        MenuCmds(i).BackImage = MenuCmds(UBound(MenuCmds)).BackImage
                        MenuCmds(i).hBackColor = MenuCmds(UBound(MenuCmds)).hBackColor
                        MenuCmds(i).HoverFont = MenuCmds(UBound(MenuCmds)).HoverFont
                        MenuCmds(i).hTextColor = MenuCmds(UBound(MenuCmds)).hTextColor
                        MenuCmds(i).iCursor = MenuCmds(UBound(MenuCmds)).iCursor
                        MenuCmds(i).LeftImage = MenuCmds(UBound(MenuCmds)).LeftImage
                        MenuCmds(i).nBackColor = MenuCmds(UBound(MenuCmds)).nBackColor
                        MenuCmds(i).NormalFont = MenuCmds(UBound(MenuCmds)).NormalFont
                        MenuCmds(i).nTextColor = MenuCmds(UBound(MenuCmds)).nTextColor
                        MenuCmds(i).RightImage = MenuCmds(UBound(MenuCmds)).RightImage
                    Next i
                    DoneStatus = DoneStatus + 1
                End If
            End If
        Loop
    Close #ff
    
    If cmbGrps.ListIndex > 0 Or cmbCmds.ListIndex > 0 Then
        If Not frmMain.tvMenus.SelectedItem.Parent Is Nothing Then
            frmMain.tvMenus.Nodes(frmMain.tvMenus.SelectedItem.Parent.Index).Selected = True
        End If
        frmMain.tvMenus.Nodes.Remove frmMain.tvMenus.SelectedItem.Index
        ReDim Preserve MenuGrps(UBound(MenuGrps) - 1)
        ReDim Preserve MenuCmds(UBound(MenuCmds) - 1)
    End If
    
    Unload Me
    
    Exit Sub
    
ReportError:
    MsgBox "An unexpected error has occurred while importing the project" + vbCrLf + "Error " & Err.Number & ": " + Err.Description, vbCritical + vbOKOnly, GetLocalizedStr(662)
    
    On Error Resume Next
    Close #ff
    Unload Me

End Sub

Private Sub Form_Load()

    Dim i As Integer
    Dim sLine As String
    
    CenterForm Me

    SrcProject = GetProjectProperties(ImportProjectFileName)
    
    Caption = Caption + " - " + SrcProject.Name
    
    For i = 0 To UBound(SrcProject.UserConfigs)
        lstConfigs.AddItem SrcProject.UserConfigs(i).Name
        lstConfigs.Selected(lstConfigs.NewIndex) = (SrcProject.DefaultConfig = i)
    Next i
    
    cmbGrps.AddItem GetLocalizedStr(110)
    cmbCmds.AddItem GetLocalizedStr(110)
    
    Do Until (LOF(ff) = Loc(ff))
        Line Input #ff, sLine
        If Left(sLine, 3) = "[G]" Then
            cmbGrps.AddItem Mid(GetParam(sLine, 1), 4)
        ElseIf Left(sLine, 5) = "[C]  " Then
            If Mid(GetParam(sLine, 1), 6) <> "[SEP]" Then
                cmbCmds.AddItem Mid(GetParam(sLine, 1), 6)
            End If
        End If
    Loop
    Close #ff
    
    cmbGrps.ListIndex = 0
    cmbCmds.ListIndex = 0

End Sub

Private Sub lstConfigs_Click()

    If IsUpdating Then Exit Sub
    lstConfigs_ItemCheck lstConfigs.ListIndex

End Sub

Private Sub lstConfigs_ItemCheck(item As Integer)

    Dim i As Integer
    
    If IsUpdating Then Exit Sub
    IsUpdating = True
    
    lstConfigs.Selected(item) = Not lstConfigs.Selected(item)
    
    IsUpdating = False

End Sub
