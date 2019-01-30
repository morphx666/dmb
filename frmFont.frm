VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmFont 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font"
   ClientHeight    =   5310
   ClientLeft      =   2325
   ClientTop       =   4020
   ClientWidth     =   6030
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFont.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5310
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frameHover 
      Caption         =   "Mouse Over"
      Height          =   2085
      Left            =   3075
      TabIndex        =   5
      Top             =   555
      Width           =   2865
      Begin VB.TextBox txtBlur2 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   1
         Left            =   2415
         TabIndex        =   35
         Text            =   "00"
         Top             =   1215
         Width           =   315
      End
      Begin VB.TextBox txtOffsetY2 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   1
         Left            =   2010
         TabIndex        =   34
         Text            =   "00"
         Top             =   1215
         Width           =   315
      End
      Begin VB.TextBox txtOffsetX2 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   33
         Text            =   "00"
         Top             =   1215
         Width           =   315
      End
      Begin VB.CheckBox chkShadow2 
         Caption         =   "Shadow 2"
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   31
         Top             =   1230
         Width           =   1140
      End
      Begin VB.TextBox txtBlur1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   1
         Left            =   2415
         TabIndex        =   30
         Text            =   "00"
         Top             =   870
         Width           =   315
      End
      Begin VB.TextBox txtOffsetY1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   1
         Left            =   2010
         TabIndex        =   29
         Text            =   "00"
         Top             =   870
         Width           =   315
      End
      Begin VB.TextBox txtOffsetX1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   1
         Left            =   1680
         TabIndex        =   28
         Text            =   "00"
         Top             =   870
         Width           =   315
      End
      Begin VB.CheckBox chkShadow1 
         Caption         =   "Shadow 1"
         Height          =   210
         Index           =   1
         Left            =   90
         TabIndex        =   26
         Top             =   885
         Width           =   1140
      End
      Begin SmartButtonProject.SmartButton cmdDefault 
         Height          =   300
         Index           =   1
         Left            =   1635
         TabIndex        =   9
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         Caption         =   "Reset"
         Picture         =   "frmFont.frx":014A
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
      Begin SmartButtonProject.SmartButton cmdChange 
         Height          =   300
         Index           =   1
         Left            =   105
         TabIndex        =   8
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         Caption         =   "Change"
         Picture         =   "frmFont.frx":04E4
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
      Begin xfxLine3D.ucLine3D uc3DLine2 
         Height          =   30
         Left            =   105
         TabIndex        =   7
         Top             =   1560
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   53
      End
      Begin SmartButtonProject.SmartButton cmdColor1 
         Height          =   240
         Index           =   1
         Left            =   1290
         TabIndex        =   27
         Top             =   885
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   6
         Enabled         =   0   'False
         ShowFocus       =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdColor2 
         Height          =   240
         Index           =   1
         Left            =   1290
         TabIndex        =   32
         Top             =   1230
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   6
         Enabled         =   0   'False
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblFont 
         BackStyle       =   0  'Transparent
         Height          =   435
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   300
         Width           =   2640
      End
   End
   Begin SmartButtonProject.SmartButton sbApplyOptions 
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   135
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   609
      Caption         =   "Options"
      Picture         =   "frmFont.frx":087E
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
   Begin MSComctlLib.ImageList ilIcons 
      Left            =   675
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFont.frx":09D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFont.frx":0AEA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmFont.frx":0BFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageCombo cmbAlign 
      Height          =   330
      Left            =   120
      TabIndex        =   11
      Top             =   3030
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   582
      _Version        =   393216
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Locked          =   -1  'True
      Text            =   "Left"
      ImageList       =   "ilIcons"
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
      TabIndex        =   12
      Top             =   3420
      Width           =   5805
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
      Left            =   5025
      TabIndex        =   14
      Top             =   4815
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
      Left            =   3975
      TabIndex        =   13
      Top             =   4815
      Width           =   900
   End
   Begin VB.Frame frameNormal 
      Caption         =   "Normal"
      Height          =   2085
      Left            =   120
      TabIndex        =   0
      Top             =   555
      Width           =   2865
      Begin VB.TextBox txtBlur2 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   0
         Left            =   2430
         TabIndex        =   25
         Text            =   "00"
         Top             =   1185
         Width           =   315
      End
      Begin VB.TextBox txtOffsetY2 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   0
         Left            =   2025
         TabIndex        =   24
         Text            =   "00"
         Top             =   1185
         Width           =   315
      End
      Begin VB.TextBox txtOffsetX2 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   0
         Left            =   1695
         TabIndex        =   23
         Text            =   "00"
         Top             =   1185
         Width           =   315
      End
      Begin VB.CheckBox chkShadow2 
         Caption         =   "Shadow 2"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   21
         Top             =   1200
         Width           =   1140
      End
      Begin VB.TextBox txtBlur1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   0
         Left            =   2430
         TabIndex        =   20
         Text            =   "00"
         Top             =   833
         Width           =   315
      End
      Begin VB.TextBox txtOffsetY1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   0
         Left            =   2025
         TabIndex        =   19
         Text            =   "00"
         Top             =   833
         Width           =   315
      End
      Begin VB.TextBox txtOffsetX1 
         Alignment       =   1  'Right Justify
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
         Height          =   285
         Index           =   0
         Left            =   1695
         TabIndex        =   18
         Text            =   "00"
         Top             =   833
         Width           =   315
      End
      Begin VB.CheckBox chkShadow1 
         Caption         =   "Shadow 1"
         Height          =   210
         Index           =   0
         Left            =   105
         TabIndex        =   16
         Top             =   855
         Width           =   1140
      End
      Begin xfxLine3D.ucLine3D uc3DLine1 
         Height          =   30
         Left            =   105
         TabIndex        =   2
         Top             =   1560
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   53
      End
      Begin SmartButtonProject.SmartButton cmdChange 
         Height          =   300
         Index           =   0
         Left            =   105
         TabIndex        =   3
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         Caption         =   "Change"
         Picture         =   "frmFont.frx":0D0E
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
      Begin SmartButtonProject.SmartButton cmdDefault 
         Height          =   300
         Index           =   0
         Left            =   1620
         TabIndex        =   4
         Top             =   1650
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   529
         Caption         =   "Reset"
         Picture         =   "frmFont.frx":10A8
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
      Begin SmartButtonProject.SmartButton cmdColor1 
         Height          =   240
         Index           =   0
         Left            =   1305
         TabIndex        =   17
         Top             =   855
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   6
         Enabled         =   0   'False
         ShowFocus       =   -1  'True
      End
      Begin SmartButtonProject.SmartButton cmdColor2 
         Height          =   240
         Index           =   0
         Left            =   1305
         TabIndex        =   22
         Top             =   1200
         Width           =   270
         _ExtentX        =   476
         _ExtentY        =   423
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureLayout   =   6
         Enabled         =   0   'False
         ShowFocus       =   -1  'True
      End
      Begin VB.Label lblFont 
         BackStyle       =   0  'Transparent
         Height          =   435
         Index           =   0
         Left            =   105
         TabIndex        =   1
         Top             =   315
         Width           =   2640
      End
   End
   Begin VB.Label lblAlignment 
      AutoSize        =   -1  'True
      Caption         =   "Alignment"
      Height          =   210
      Left            =   120
      TabIndex        =   10
      Top             =   2805
      Width           =   825
   End
End
Attribute VB_Name = "frmFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim NewFont(0 To 1) As SelFontDef
Dim BackCmd As MenuCmd

Private IsLoading As Boolean

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub ShowWarning()

    With TipsSys
        Do While .IsVisible
            DoEvents
        Loop
        .TipTitle = "Text Shadow"
        .Tip = "Text shadow effect is only available in FireFox, Safari Opera and Chrome. Since Internet Explorer does not support it, the preview will not reflect your changes."
        .CanDisable = True
        .Show
    End With

End Sub

Private Sub chkShadow1_Click(Index As Integer)

    ShowWarning

    cmdColor1(Index).Enabled = chkShadow1(Index).Value
    txtOffsetX1(Index).Enabled = chkShadow1(Index).Value
    txtOffsetY1(Index).Enabled = chkShadow1(Index).Value
    txtBlur1(Index).Enabled = chkShadow1(Index).Value
    
    UpdateFont Index

End Sub

Private Sub chkShadow2_Click(Index As Integer)

    ShowWarning

    cmdColor2(Index).Enabled = chkShadow2(Index).Value
    txtOffsetX2(Index).Enabled = chkShadow2(Index).Value
    txtOffsetY2(Index).Enabled = chkShadow2(Index).Value
    txtBlur2(Index).Enabled = chkShadow2(Index).Value
    
    UpdateFont Index

End Sub

Private Sub cmbAlign_Click()

    MenuCmds(GetID).Alignment = cmbAlign.SelectedItem.Index - 1
    frmMain.DoLivePreview wbLivePreview

End Sub

Private Sub cmdCancel_Click()

    MenuCmds(GetID) = BackCmd
    Unload Me

End Sub

Private Sub cmdChange_Click(Index As Integer)

    SelFont = NewFont(Index)

    With frmFontDialog
        .Show vbModal
        CenterForm frmFontDialog
        Do
            DoEvents
        Loop While .Visible
    End With
    
    UpdateFont Index

End Sub

Private Sub UpdateFont(Index As Integer)

    If IsLoading Then Exit Sub

    With NewFont(Index).Shadow
        .Enabled1 = chkShadow1(Index).Value = vbChecked
        .Color1 = cmdColor1(Index).BackColor
        .OffsetX1 = Val(txtOffsetX1(Index).Text)
        .OffsetY1 = Val(txtOffsetY1(Index).Text)
        .Blur1 = Val(txtBlur1(Index).Text)
        
        .Enabled2 = chkShadow2(Index).Value = vbChecked
        .Color2 = cmdColor2(Index).BackColor
        .OffsetX2 = Val(txtOffsetX2(Index).Text)
        .OffsetY2 = Val(txtOffsetY2(Index).Text)
        .Blur2 = Val(txtBlur2(Index).Text)
    End With
    
    With MenuCmds(GetID)
        If Index = 0 Then
            .NormalFont.FontShadow.Enabled1 = NewFont(0).Shadow.Enabled1
            .NormalFont.FontShadow.Color1 = NewFont(0).Shadow.Color1
            .NormalFont.FontShadow.OffsetX1 = NewFont(0).Shadow.OffsetX1
            .NormalFont.FontShadow.OffsetY1 = NewFont(0).Shadow.OffsetY1
            .NormalFont.FontShadow.Blur1 = NewFont(0).Shadow.Blur1
            .NormalFont.FontShadow.Enabled2 = NewFont(0).Shadow.Enabled2
            .NormalFont.FontShadow.Color2 = NewFont(0).Shadow.Color2
            .NormalFont.FontShadow.OffsetX2 = NewFont(0).Shadow.OffsetX2
            .NormalFont.FontShadow.OffsetY2 = NewFont(0).Shadow.OffsetY2
            .NormalFont.FontShadow.Blur2 = NewFont(0).Shadow.Blur2
        End If
        
        If Index = 1 Then
            .HoverFont.FontShadow.Enabled1 = NewFont(1).Shadow.Enabled1
            .HoverFont.FontShadow.Color1 = NewFont(1).Shadow.Color1
            .HoverFont.FontShadow.OffsetX1 = NewFont(1).Shadow.OffsetX1
            .HoverFont.FontShadow.OffsetY1 = NewFont(1).Shadow.OffsetY1
            .HoverFont.FontShadow.Blur1 = NewFont(1).Shadow.Blur1
            .HoverFont.FontShadow.Enabled2 = NewFont(1).Shadow.Enabled2
            .HoverFont.FontShadow.Color2 = NewFont(1).Shadow.Color2
            .HoverFont.FontShadow.OffsetX2 = NewFont(1).Shadow.OffsetX2
            .HoverFont.FontShadow.OffsetY2 = NewFont(1).Shadow.OffsetY2
            .HoverFont.FontShadow.Blur2 = NewFont(1).Shadow.Blur2
        End If
    End With

    If SelFont.IsValid Then
        NewFont(Index) = SelFont
        
        With MenuCmds(GetID)
            .NormalFont.FontName = NewFont(0).Name
            .NormalFont.FontSize = NewFont(0).Size
            .NormalFont.FontBold = NewFont(0).Bold
            .NormalFont.FontItalic = NewFont(0).Italic
            .NormalFont.FontUnderline = NewFont(0).Underline
                       
            .HoverFont.FontName = NewFont(1).Name
            .HoverFont.FontSize = NewFont(1).Size
            .HoverFont.FontBold = NewFont(1).Bold
            .HoverFont.FontItalic = NewFont(1).Italic
            .HoverFont.FontUnderline = NewFont(1).Underline
        End With
        
        UpdateFontLabel Index
    End If
    
    frmMain.DoLivePreview wbLivePreview

End Sub

Private Sub UpdateFontObject()

    With MenuCmds(GetID)
        caption = NiceGrpCaption(.parent) + "/" + NiceCmdCaption(GetID) + " - " + GetLocalizedStr(213)
        NewFont(0).Name = .NormalFont.FontName
        NewFont(0).Size = .NormalFont.FontSize
        NewFont(0).Bold = .NormalFont.FontBold
        NewFont(0).Italic = .NormalFont.FontItalic
        NewFont(0).Underline = .NormalFont.FontUnderline
        
        NewFont(0).Shadow.Enabled1 = .NormalFont.FontShadow.Enabled1
        NewFont(0).Shadow.Color1 = .NormalFont.FontShadow.Color1
        NewFont(0).Shadow.OffsetX1 = .NormalFont.FontShadow.OffsetX1
        NewFont(0).Shadow.OffsetY1 = .NormalFont.FontShadow.OffsetY1
        NewFont(0).Shadow.Blur1 = .NormalFont.FontShadow.Blur1
        NewFont(0).Shadow.Enabled2 = .NormalFont.FontShadow.Enabled2
        NewFont(0).Shadow.Color2 = .NormalFont.FontShadow.Color2
        NewFont(0).Shadow.OffsetX2 = .NormalFont.FontShadow.OffsetX2
        NewFont(0).Shadow.OffsetY2 = .NormalFont.FontShadow.OffsetY2
        NewFont(0).Shadow.Blur2 = .NormalFont.FontShadow.Blur2
        
        NewFont(1).Name = .HoverFont.FontName
        NewFont(1).Size = .HoverFont.FontSize
        NewFont(1).Bold = .HoverFont.FontBold
        NewFont(1).Italic = .HoverFont.FontItalic
        NewFont(1).Underline = .HoverFont.FontUnderline
        
        NewFont(1).Shadow.Enabled1 = .HoverFont.FontShadow.Enabled1
        NewFont(1).Shadow.Color1 = .HoverFont.FontShadow.Color1
        NewFont(1).Shadow.OffsetX1 = .HoverFont.FontShadow.OffsetX1
        NewFont(1).Shadow.OffsetY1 = .HoverFont.FontShadow.OffsetY1
        NewFont(1).Shadow.Blur1 = .HoverFont.FontShadow.Blur1
        NewFont(1).Shadow.Enabled2 = .HoverFont.FontShadow.Enabled2
        NewFont(1).Shadow.Color2 = .HoverFont.FontShadow.Color2
        NewFont(1).Shadow.OffsetX2 = .HoverFont.FontShadow.OffsetX2
        NewFont(1).Shadow.OffsetY2 = .HoverFont.FontShadow.OffsetY2
        NewFont(1).Shadow.Blur2 = .HoverFont.FontShadow.Blur2
    End With

End Sub

Private Sub cmdColor1_Click(Index As Integer)

    BuildUsedColorsArray

    With cmdColor1(Index)
        SelColor = .BackColor
        SelColor_CanBeTransparent = False
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdColor1(Index)
    End With
    
    cmdColor1(Index).ZOrder 0
    
    UpdateFont Index

End Sub

Private Sub cmdColor2_Click(Index As Integer)

    BuildUsedColorsArray

    With cmdColor2(Index)
        SelColor = .BackColor
        SelColor_CanBeTransparent = False
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdColor2(Index)
    End With
    
    cmdColor2(Index).ZOrder 0
    
    UpdateFont Index

End Sub

Private Sub cmdDefault_Click(Index As Integer)

    Select Case Index
        Case 0
            MenuCmds(GetID).NormalFont = MenuGrps(MenuCmds(GetID).parent).DefNormalFont
        Case 1
            MenuCmds(GetID).HoverFont = MenuGrps(MenuCmds(GetID).parent).DefHoverFont
    End Select
    
    UpdateFontObject
    frmMain.DoLivePreview wbLivePreview
    UpdateFontLabel Index

End Sub

Private Sub cmdOK_Click()

    ApplyStyleOptions
    frmMain.SaveState "Change " + MenuCmds(GetID).Name + " Font"
    
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
        .HoverFont = MenuCmds(sId).HoverFont
        .NormalFont = MenuCmds(sId).NormalFont
        .Alignment = MenuCmds(sId).Alignment
    End With

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/command_font.htm"

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
    
    BackCmd = MenuCmds(GetID)
    
    cmbAlign.ComboItems.Add , , GetLocalizedStr(190), 1
    cmbAlign.ComboItems.Add , , GetLocalizedStr(191), 3
    cmbAlign.ComboItems.Add , , GetLocalizedStr(192), 2
    cmbAlign.ComboItems(MenuCmds(GetID).Alignment + 1).Selected = True
    
    UpdateFontObject
    UpdateFontLabel 0
    UpdateFontLabel 1
        
End Sub

Private Sub UpdateFontLabel(Index As Integer)

    On Error Resume Next

    With lblFont(Index)
        .caption = NewFont(Index).Name + ", "
        .caption = .caption & NewFont(Index).Size
        If NewFont(Index).Bold Then
            .caption = .caption & ", Bold"
        End If
        If NewFont(Index).Italic Then
            .caption = .caption & ", Italic"
        End If
        If NewFont(Index).Underline Then
            .caption = .caption & ", Underline"
        End If
    End With
    
    With lblFont(Index)
        .FontName = NewFont(Index).Name
        .FontBold = NewFont(Index).Bold
        .FontItalic = NewFont(Index).Italic
        .FontSize = px2pt(NewFont(Index).Size)
        .FontUnderline = NewFont(Index).Underline
    End With
    
    IsLoading = True
    
    With MenuCmds(GetID).NormalFont.FontShadow
        chkShadow1(0).Value = IIf(.Enabled1, vbChecked, vbUnchecked)
        cmdColor1(0).BackColor = .Color1
        txtOffsetX1(0).Text = .OffsetX1
        txtOffsetY1(0).Text = .OffsetY1
        txtBlur1(0).Text = .Blur1
        
        chkShadow2(0).Value = IIf(.Enabled2, vbChecked, vbUnchecked)
        cmdColor2(0).BackColor = .Color2
        txtOffsetX2(0).Text = .OffsetX2
        txtOffsetY2(0).Text = .OffsetY2
        txtBlur2(0).Text = .Blur2
    End With
    
    With MenuCmds(GetID).HoverFont.FontShadow
        chkShadow1(1).Value = IIf(.Enabled1, vbChecked, vbUnchecked)
        cmdColor1(1).BackColor = .Color1
        txtOffsetX1(1).Text = .OffsetX1
        txtOffsetY1(1).Text = .OffsetY1
        txtBlur1(1).Text = .Blur1
        
        chkShadow2(1).Value = IIf(.Enabled2, vbChecked, vbUnchecked)
        cmdColor2(1).BackColor = .Color2
        txtOffsetX2(1).Text = .OffsetX2
        txtOffsetY2(1).Text = .OffsetY2
        txtBlur2(1).Text = .Blur2
    End With
    
    IsLoading = False

End Sub

Private Sub sbApplyOptions_Click()

    PopupMenu frmMain.mnuStyleOptions, , sbApplyOptions.Left, sbApplyOptions.Top + sbApplyOptions.Height

End Sub

Private Sub LocalizeUI()

    frameNormal.caption = GetLocalizedStr(179)
    frameHover.caption = GetLocalizedStr(180)

    cmdDefault(0).caption = GetLocalizedStr(183)
    cmdDefault(1).caption = GetLocalizedStr(183)
    cmdChange(0).caption = GetLocalizedStr(189)
    cmdChange(1).caption = GetLocalizedStr(189)
    lblAlignment.caption = GetLocalizedStr(115)
    
    frmLiveSample.caption = GetLocalizedStr(188)
    
    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdDefault(0).Width = SetCtrlWidth(cmdDefault(0))
        cmdDefault(1).Width = SetCtrlWidth(cmdDefault(1))
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub

Private Sub txtBlur1_Change(Index As Integer)

    UpdateFont Index

End Sub

Private Sub txtBlur2_Change(Index As Integer)

    UpdateFont Index

End Sub

Private Sub txtOffsetX1_Change(Index As Integer)

    UpdateFont Index

End Sub

Private Sub txtOffsetX2_Change(Index As Integer)

    UpdateFont Index

End Sub

Private Sub txtOffsetY1_Change(Index As Integer)

    UpdateFont Index

End Sub

Private Sub txtOffsetY2_Change(Index As Integer)

    UpdateFont Index

End Sub
