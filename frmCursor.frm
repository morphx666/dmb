VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmCursor 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cursor"
   ClientHeight    =   4275
   ClientLeft      =   6660
   ClientTop       =   6135
   ClientWidth     =   4725
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCursor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MouseIcon       =   "frmCursor.frx":014A
   ScaleHeight     =   4275
   ScaleWidth      =   4725
   ShowInTaskbar   =   0   'False
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   135
      TabIndex        =   10
      Top             =   135
      Width           =   4470
      _ExtentX        =   7885
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmCursor.frx":029C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionLayout   =   3
      PictureLayout   =   3
   End
   Begin VB.Frame frmLiveSample 
      Caption         =   "Live Sample"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   7
      Top             =   2295
      Width           =   4500
   End
   Begin VB.CommandButton cmdCancel 
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
      Height          =   375
      Left            =   3690
      TabIndex        =   9
      Top             =   3765
      Width           =   900
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2655
      TabIndex        =   8
      Top             =   3765
      Width           =   900
   End
   Begin VB.Frame frameHover 
      Caption         =   "Mouse Over"
      Height          =   1605
      Left            =   120
      TabIndex        =   0
      Top             =   555
      Width           =   4500
      Begin VB.PictureBox picCursor 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   3495
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   5
         Top             =   945
         Width           =   480
      End
      Begin VB.ComboBox cmbCursors 
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
         ItemData        =   "frmCursor.frx":03F6
         Left            =   2145
         List            =   "frmCursor.frx":03F8
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   390
         Width           =   1845
      End
      Begin SmartButtonProject.SmartButton cmdChange 
         Height          =   240
         Left            =   2145
         TabIndex        =   3
         Top             =   945
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Change"
         Picture         =   "frmCursor.frx":03FA
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin SmartButtonProject.SmartButton cmdRemove 
         Height          =   240
         Left            =   2145
         TabIndex        =   6
         Top             =   1185
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   423
         Caption         =   "Remove"
         Picture         =   "frmCursor.frx":0794
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CaptionLayout   =   4
         PictureLayout   =   3
      End
      Begin VB.Label lblCrFileName 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Custom Cursor Icon"
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
         Left            =   555
         TabIndex        =   4
         Top             =   1095
         Width           =   1425
      End
      Begin VB.Label lblCrType 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cursor Type"
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
         Left            =   1095
         TabIndex        =   1
         Top             =   450
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim BackCmd As MenuCmd
Dim IsUpdating As Boolean

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub cmdCancel_Click()

    MenuCmds(GetID) = BackCmd
    Unload Me

End Sub

Private Sub cmbCursors_Click()

    cmdChange.Enabled = (cmbCursors.ListIndex = (cmbCursors.ListCount - 1))
    cmdRemove.Enabled = cmdChange.Enabled
    lblCrFileName.Enabled = cmdChange.Enabled

    If IsUpdating Then Exit Sub

    If cmbCursors.ListIndex <= 12 Then
        MenuCmds(GetID).iCursor.cType = cmbCursors.ListIndex + 1
    Else
        Select Case cmbCursors.ListIndex
            Case 13
                MenuCmds(GetID).iCursor.cType = iccResizeAll
            Case 14
                MenuCmds(GetID).iCursor.cType = iccWait
            Case 15
                MenuCmds(GetID).iCursor.cType = iccCustom
        End Select
    End If
    frmMain.DoLivePreview wbLivePreview

End Sub

Private Sub cmdOK_Click()

    ApplyStyleOptions
    frmMain.SaveState "Change " + MenuCmds(GetID).Name + " Cursor"
    
    Unload Me

End Sub

Private Sub ApplyStyleOptions()

    Dim i As Integer
    Dim c As Integer
    Dim t As Integer
    Dim sId As Integer
    
    sId = GetID
    
    For c = 0 To frmMain.mnuStyleOptionsOP.Count - 1
        If frmMain.mnuStyleOptionsOP(c).Checked Then
            t = Val(frmMain.mnuStyleOptionsOP(c).tag)
            Select Case c
                Case 0: ' do nothing
                Case 1:
                    For i = 1 To UBound(MenuCmds)
                        If MenuCmds(i).parent = t Then CopyStyle sId, i
                    Next i
                Case 2:
                    For i = 1 To UBound(MenuCmds)
                        If BelongsToToolbar(i, False) = t Then CopyStyle sId, i
                    Next i
                Case 3:
                    For i = 1 To UBound(MenuCmds)
                        CopyStyle sId, i
                    Next i
            End Select
            Exit Sub
        End If
    Next c
    
    With dmbClipboard
        For i = 1 To UBound(.CustomSel)
            CopyStyle sId, GetIDByName(.CustomSel(i))
        Next i
    End With

End Sub

Private Sub CopyStyle(sId As Integer, tID As Integer)

    With MenuCmds(tID)
        .iCursor = MenuCmds(sId).iCursor
    End With

End Sub

Private Sub cmdRemove_Click()

    MenuCmds(GetID).iCursor.CFile = ""

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/group_and_command_cursor.htm"

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd Me.hwnd, False

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    Select Case hwnd
        Case Me.hwnd
            Select Case uMsg
                Case WM_CTLCOLORSTATIC, WM_CTLCOLOREDIT, WM_PAINT
                    DrawColorBoxes Me
            End Select
    End Select
    
End Sub

Private Sub Form_Load()

    CenterForm Me
    LocalizeUI
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True
    
    cmbCursors.AddItem GetLocalizedStr(193)
    cmbCursors.AddItem GetLocalizedStr(194)
    cmbCursors.AddItem GetLocalizedStr(195)
    cmbCursors.AddItem GetLocalizedStr(196)
    cmbCursors.AddItem GetLocalizedStr(197)
    cmbCursors.AddItem "Size E"
    cmbCursors.AddItem "Size NE"
    cmbCursors.AddItem "Size NW"
    cmbCursors.AddItem "Size N"
    cmbCursors.AddItem "Size SE"
    cmbCursors.AddItem "Size SW"
    cmbCursors.AddItem "Size S"
    cmbCursors.AddItem "Size W"
    cmbCursors.AddItem "Size All"
    cmbCursors.AddItem "Wait"
    cmbCursors.AddItem "Custom"
    
    BackCmd = MenuCmds(GetID)
    
    IsUpdating = True
    
    cmbCursors.ListIndex = 0
    With MenuCmds(GetID)
        Select Case .iCursor.cType
            Case iccDefault
                cmbCursors.ListIndex = 0
            Case iccCrosshair
                cmbCursors.ListIndex = 1
            Case iccHand
                cmbCursors.ListIndex = 2
            Case iccText
                cmbCursors.ListIndex = 3
            Case iccHelp
                cmbCursors.ListIndex = 4
            Case iccResizeE
                cmbCursors.ListIndex = 5
            Case iccResizeNE
                cmbCursors.ListIndex = 6
            Case iccResizeNW
                cmbCursors.ListIndex = 7
            Case iccResizeN
                cmbCursors.ListIndex = 8
            Case iccResizeSE
                cmbCursors.ListIndex = 9
            Case iccResizeSW
                cmbCursors.ListIndex = 10
            Case iccResizeS
                cmbCursors.ListIndex = 11
            Case iccResizeW
                cmbCursors.ListIndex = 12
            Case iccResizeAll
                cmbCursors.ListIndex = 13
            Case iccWait
                cmbCursors.ListIndex = 14
            Case iccCustom
                cmbCursors.ListIndex = 15
        End Select
        
        caption = NiceGrpCaption(.parent) + "/" + NiceCmdCaption(GetID) + " - " + GetLocalizedStr(215)
    End With
    
    IsUpdating = False
    
    UpdateCursorImage True

End Sub

Private Sub cmdChange_Click()

    With SelImage
        .FileName = picCursor.tag
        .LimitToCursors = True
        
        frmRscImages.Show vbModal
        
        If .IsValid Then
            MenuCmds(GetID).iCursor.CFile = .FileName
            UpdateCursorImage
        End If
    End With

End Sub

Private Sub UpdateCursorImage(Optional IsLoading As Boolean)

    If IsANI(MenuCmds(GetID).iCursor.CFile) Then
        Set picCursor.Picture = LoadResPicture(201, vbResIcon)
    Else
        Set picCursor.Picture = LoadPictureRes(MenuCmds(GetID).iCursor.CFile)
    End If

    If Not IsLoading Then frmMain.DoLivePreview wbLivePreview

End Sub

Private Sub sbApplyOptions_Click()
    
    PopupMenu frmMain.mnuStyleOptions, , sbApplyOptions.Left, sbApplyOptions.Top + sbApplyOptions.Height

End Sub

Private Sub LocalizeUI()

    frameHover.caption = GetLocalizedStr(180)

    frmLiveSample.caption = GetLocalizedStr(188)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
