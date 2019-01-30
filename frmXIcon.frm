VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Object = "{32587D54-A348-4594-BB8D-D1DE4B265074}#1.0#0"; "IContainer.ocx"
Begin VB.Form frmXIcon 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Extract Icon"
   ClientHeight    =   4320
   ClientLeft      =   6045
   ClientTop       =   4965
   ClientWidth     =   5745
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIconSrc2 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   2910
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   3660
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.OptionButton opSize 
      Caption         =   "32x32"
      Height          =   225
      Index           =   1
      Left            =   4695
      TabIndex        =   11
      Top             =   2820
      Width           =   840
   End
   Begin VB.OptionButton opSize 
      Caption         =   "16x16"
      Height          =   225
      Index           =   0
      Left            =   4695
      TabIndex        =   10
      Top             =   2595
      Value           =   -1  'True
      Width           =   840
   End
   Begin MSComDlg.CommonDialog cDlg 
      Left            =   1425
      Top             =   3585
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DefaultExt      =   "0"
      DialogTitle     =   "Select Project"
      Filter          =   "DHTML Menu Builder Projects|*.dmb"
      MaxFileSize     =   1024
   End
   Begin VB.Timer tmrR 
      Interval        =   250
      Left            =   690
      Top             =   3675
   End
   Begin VB.PictureBox picIconSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   3585
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3675
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   4822
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1020
      Width           =   540
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4515
      TabIndex        =   16
      Top             =   3900
      Width           =   1155
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save..."
      Default         =   -1  'True
      Height          =   360
      Left            =   4515
      TabIndex        =   12
      Top             =   3360
      Width           =   1155
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   -30
      TabIndex        =   15
      Top             =   3795
      Width           =   5970
      _ExtentX        =   10530
      _ExtentY        =   53
   End
   Begin IContainer.IconContainer icIcons 
      Height          =   2745
      Left            =   45
      TabIndex        =   5
      Top             =   1020
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4842
   End
   Begin VB.TextBox txtFileName 
      Height          =   285
      Left            =   60
      TabIndex        =   1
      Top             =   315
      Width           =   3840
   End
   Begin SmartButtonProject.SmartButton cmdBrowse 
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   300
      Width           =   360
      _ExtentX        =   635
      _ExtentY        =   556
      Picture         =   "frmXIcon.frx":0000
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
   Begin SmartButtonProject.SmartButton cmdColor 
      Height          =   240
      Left            =   4957
      TabIndex        =   8
      Top             =   1935
      Width           =   270
      _ExtentX        =   476
      _ExtentY        =   423
      BackColor       =   16777215
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocus       =   -1  'True
   End
   Begin VB.Label lblIconSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Size"
      Height          =   195
      Left            =   4770
      TabIndex        =   9
      Top             =   2340
      Width           =   645
   End
   Begin VB.Label lblBackColorN 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Back Color"
      Height          =   195
      Left            =   4717
      TabIndex        =   7
      Top             =   1695
      Width           =   750
   End
   Begin VB.Label lblSelIcon 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Selected Icon"
      Height          =   195
      Left            =   4605
      TabIndex        =   3
      Top             =   765
      Width           =   975
   End
   Begin VB.Label lblLoadedIcons 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loaded Icons"
      Height          =   195
      Left            =   60
      TabIndex        =   4
      Top             =   795
      Width           =   960
   End
   Begin VB.Label lblIconLibrary 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Icon Library"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   75
      Width           =   855
   End
End
Attribute VB_Name = "frmXIcon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents cGif As GIF
Attribute cGif.VB_VarHelpID = -1

Private Sub cmdSave_Click()

    On Error GoTo ExitSub
    
    If TypeOf ActiveControl Is TextBox Then
        If ActiveControl.Name = txtFileName.Name Then
            LoadIconLib
            Exit Sub
        End If
    End If

    With cDlg
        .CancelError = True
        .DialogTitle = "Select Location to Save Icon"
        .Flags = cdlOFNPathMustExist + cdlOFNHideReadOnly + cdlOFNLongNames + cdlOFNNoReadOnlyReturn + cdlOFNOverwritePrompt
        .filter = "CompuServe GIF|*.gif"
        .FileName = GetFileName(txtFileName.Text)
        .FileName = Left(.FileName, Len(.FileName) - 4) + ".gif"
        .ShowSave
        
        SaveAsGIF .FileName
    End With
    
ExitSub:

End Sub

Private Sub SaveAsGIF(FileName As String)

    Set cGif = New GIF

    If cmdColor.Tag = -2 Then
        SetColor &H80, cmdColor
        icIcons_IconSelected icIcons.selectedIndex
        cGif.SaveGIF picIconSrc.Image, FileName, picIconSrc.hDc, True, cmdColor.BackColor
        SetColor -2, cmdColor
        icIcons_IconSelected icIcons.selectedIndex
    Else
        cGif.SaveGIF picIconSrc.Image, FileName, picIconSrc.hDc, False
    End If
    
    Set cGif = Nothing

End Sub

Private Sub cmdBrowse_Click()

    With cDlg
        .DialogTitle = "Select Icon Library"
        .Flags = cdlOFNCreatePrompt + cdlOFNHideReadOnly + cdlOFNLongNames
        .filter = "Icon Libraries|*.dll;*.exe;*.icl;*.cpl|All Files|*.*"
        .FileName = ""
        On Error Resume Next
        .ShowOpen
        On Error GoTo 0
        If Err.number > 0 Or LenB(.FileName) = 0 Then
            Exit Sub
        End If
        txtFileName.Text = .FileName
    End With

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdColor_Click()

    BuildUsedColorsArray

    SelColor = cmdColor.Tag
    SelColor_CanBeTransparent = True
    frmColorPicker.Show vbModal
    cmdColor.Tag = SelColor
    SetColor SelColor, cmdColor
    icIcons_IconSelected icIcons.selectedIndex

End Sub

Private Sub Form_Load()

    DisplayTip "Copyright Warning", "The use of icons from some applications may be in violation of their copyright." + vbCrLf + vbCrLf + "Make sure you consult the application's documentation and check if you're allowed to use its icons", True
    
    icIcons.Style = icSmall
    SetColor -2, cmdColor

    CenterForm Me
    SetupCharset Me
    LocalizeUI

End Sub

Private Sub icIcons_IconSelected(iconIdx As Long)

    Dim s As Integer

    picIcon.Cls
    picIcon.Picture = LoadPicture
    picIconSrc.Cls
    picIconSrc.Picture = LoadPicture
    picIconSrc.BackColor = vbWhite
    picIconSrc2.Cls
    picIconSrc2.Picture = LoadPicture
    picIconSrc2.BackColor = vbWhite
    
    If opSize(0).Value Then
        picIconSrc.Width = 240
        picIconSrc.Height = 240
        picIconSrc2.Width = 240
        picIconSrc2.Height = 240
    Else
        picIconSrc.Width = 480
        picIconSrc.Height = 480
        picIconSrc2.Width = 480
        picIconSrc2.Height = 480
    End If
    
    s = (picIcon.Width - picIconSrc.Width) / 2 - 15
    
    If cmdColor.Tag = -2 Then
        icIcons.GetIcon picIconSrc, iconIdx, IIf(opSize(0).Value, icSmall, icLarge), True
        picIconSrc2.PaintPicture picIconSrc.Image, 0, 0, , , , , , , vbNotSrcCopy
        
        icIcons.GetIcon picIconSrc, iconIdx, IIf(opSize(0).Value, icSmall, icLarge)
        picIconSrc2.PaintPicture picIconSrc.Image, 0, 0, , , , , , , vbNotSrcCopy
        
        SetColor cmdColor.Tag, picIcon
        icIcons.GetIcon picIconSrc, iconIdx, IIf(opSize(0).Value, icSmall, icLarge), True
        picIcon.PaintPicture picIconSrc.Image, 0, 0, , , , , , , vbSrcAnd
        picIconSrc.PaintPicture picIconSrc2.Image, 0, 0, , , , , , , vbSrcInvert
        picIcon.PaintPicture picIconSrc.Image, 0, 0, , , , , , , vbSrcInvert
        
        picIconSrc.PaintPicture picIcon.Image, 0, 0, , , , , , , vbNotSrcCopy
        picIcon.Cls
        picIcon.PaintPicture picIconSrc.Image, s, s, , , , , , , vbSrcCopy
    Else
        SetColor cmdColor.Tag, picIconSrc
        icIcons.GetIcon picIconSrc, iconIdx, IIf(opSize(0).Value, icSmall, icLarge)
    
        picIcon.PaintPicture picIconSrc.Image, s, s
    End If

End Sub

Private Sub icIcons_LoadComplete(icons As Long)

    icIcons_IconSelected icIcons.selectedIndex

End Sub

Private Sub opSize_Click(Index As Integer)

    icIcons.Style = IIf(opSize(0).Value, icSmall, icLarge)

    icIcons_IconSelected icIcons.selectedIndex

End Sub

Private Sub tmrR_Timer()

    DrawColorBoxes Me

End Sub

Private Sub txtFileName_Change()

    If FileExists(txtFileName.Text) Then LoadIconLib

End Sub

Private Sub LoadIconLib()

    icIcons.UnloadLibrary
    icIcons.IconLibrary = txtFileName.Text

End Sub

Private Sub LocalizeUI()

    Caption = GetLocalizedStr(835)
    lblIconLibrary.Caption = GetLocalizedStr(880)
    lblLoadedIcons.Caption = GetLocalizedStr(881)
    lblSelIcon.Caption = GetLocalizedStr(882)
    lblBackColorN.Caption = GetLocalizedStr(182)
    lblIconSize.Caption = GetLocalizedStr(883)
    cmdSave.Caption = GetLocalizedStr(130)
    cmdClose.Caption = GetLocalizedStr(424)

End Sub
