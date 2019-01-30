VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Object = "{C2F46FB4-62DF-499A-9E3D-EABC8CE04899}#72.0#0"; "SmartViewport.ocx"
Begin VB.Form frmStyleContainerStyle 
   ClientHeight    =   6150
   ClientLeft      =   4110
   ClientTop       =   5385
   ClientWidth     =   8805
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
   ScaleHeight     =   6150
   ScaleWidth      =   8805
   Begin SmartViewportProject.SmartViewport svpMain 
      Height          =   3180
      Left            =   570
      TabIndex        =   0
      Top             =   300
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   5609
      ScrollBarType   =   1
      ScrollLeftRight =   0   'False
      ScrollSmallChange=   300
      ButtonChange    =   300
      Begin VB.Frame frameColor 
         Caption         =   "Color"
         Height          =   2895
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   2535
         Begin VB.ComboBox cmbFX 
            Height          =   315
            ItemData        =   "frmStyleContainer.frx":0000
            Left            =   1200
            List            =   "frmStyleContainer.frx":000D
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   2385
            Width           =   1065
         End
         Begin VB.ComboBox cmbBorder 
            Height          =   315
            ItemData        =   "frmStyleContainer.frx":0029
            Left            =   1200
            List            =   "frmStyleContainer.frx":004E
            TabIndex        =   22
            Text            =   "cmbBorder"
            Top             =   1995
            WhatsThisHelpID =   20240
            Width           =   1065
         End
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   180
            Index           =   7
            Left            =   1320
            TabIndex        =   24
            Top             =   1455
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   318
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
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   240
            Index           =   0
            Left            =   1455
            TabIndex        =   25
            Top             =   285
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
            ShowFocus       =   -1  'True
         End
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   540
            Index           =   6
            Left            =   1860
            TabIndex        =   26
            Top             =   915
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   953
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
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   540
            Index           =   8
            Left            =   1140
            TabIndex        =   27
            Top             =   915
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   953
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
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   180
            Index           =   5
            Left            =   1320
            TabIndex        =   28
            Top             =   735
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   318
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
            ShowFocus       =   -1  'True
         End
         Begin xfxLine3D.ucLine3D uc3DLineColor 
            Height          =   30
            Left            =   45
            TabIndex        =   29
            Top             =   1815
            Width           =   2430
            _ExtentX        =   4286
            _ExtentY        =   53
         End
         Begin VB.Label lblBorderStyle 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Border Style"
            Height          =   195
            Left            =   120
            TabIndex        =   33
            Top             =   2445
            Width           =   885
         End
         Begin VB.Label lblBorderSize 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Border Size"
            Height          =   195
            Left            =   195
            TabIndex        =   32
            Top             =   2055
            Width           =   810
         End
         Begin VB.Label lblCorners 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corners"
            Height          =   195
            Left            =   420
            TabIndex        =   31
            Top             =   1095
            Width           =   570
         End
         Begin VB.Label lblGBackColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Back Color"
            Height          =   195
            Left            =   240
            TabIndex        =   30
            Top             =   315
            Width           =   750
         End
      End
      Begin VB.Frame frameImage 
         Caption         =   "Image"
         Height          =   2895
         Left            =   2685
         TabIndex        =   1
         Top             =   0
         Width           =   3975
         Begin VB.PictureBox picCorner 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2775
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   975
            Width           =   480
         End
         Begin VB.PictureBox picBackImage 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2775
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   2085
            Width           =   480
         End
         Begin SmartButtonProject.SmartButton sbCImage 
            Height          =   180
            Index           =   0
            Left            =   1410
            TabIndex        =   4
            Top             =   435
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   318
            BackColor       =   8421504
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
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   240
            Left            =   1410
            TabIndex        =   5
            Top             =   2085
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmStyleContainer.frx":0079
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
            Left            =   1410
            TabIndex        =   6
            Top             =   2325
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmStyleContainer.frx":0413
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
         Begin xfxLine3D.ucLine3D uc3DLineImage 
            Height          =   30
            Left            =   60
            TabIndex        =   7
            Top             =   1815
            Width           =   3870
            _ExtentX        =   6826
            _ExtentY        =   53
         End
         Begin SmartButtonProject.SmartButton sbCImage 
            Height          =   180
            Index           =   1
            Left            =   1590
            TabIndex        =   8
            Top             =   435
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   318
            BackColor       =   8388608
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
         Begin SmartButtonProject.SmartButton sbCImage 
            Height          =   180
            Index           =   2
            Left            =   2130
            TabIndex        =   9
            Top             =   435
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   318
            BackColor       =   14737632
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
         Begin SmartButtonProject.SmartButton sbCImage 
            Height          =   540
            Index           =   3
            Left            =   1410
            TabIndex        =   10
            Top             =   615
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   953
            BackColor       =   14737632
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
         Begin SmartButtonProject.SmartButton sbCImage 
            Height          =   540
            Index           =   4
            Left            =   2130
            TabIndex        =   11
            Top             =   615
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   953
            BackColor       =   14737632
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
         Begin SmartButtonProject.SmartButton sbCImage 
            Height          =   180
            Index           =   5
            Left            =   1410
            TabIndex        =   12
            Top             =   1155
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   318
            BackColor       =   14737632
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
         Begin SmartButtonProject.SmartButton sbCImage 
            Height          =   180
            Index           =   6
            Left            =   1590
            TabIndex        =   13
            Top             =   1155
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   318
            BackColor       =   14737632
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
         Begin SmartButtonProject.SmartButton sbCImage 
            Height          =   180
            Index           =   7
            Left            =   2130
            TabIndex        =   14
            Top             =   1155
            Width           =   180
            _ExtentX        =   318
            _ExtentY        =   318
            BackColor       =   14737632
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
         Begin SmartButtonProject.SmartButton sbChangeCImage 
            Height          =   240
            Left            =   2415
            TabIndex        =   15
            Top             =   435
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmStyleContainer.frx":07AD
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
         Begin SmartButtonProject.SmartButton sbRemoveCImage 
            Height          =   240
            Left            =   2415
            TabIndex        =   16
            Top             =   675
            Width           =   1215
            _ExtentX        =   2143
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmStyleContainer.frx":0B47
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
         Begin VB.Label lblCornerImages 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Corner Images"
            Height          =   195
            Left            =   105
            TabIndex        =   20
            Top             =   855
            Width           =   1065
         End
         Begin VB.Label lblSelCImage 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Top Left"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1410
            TabIndex        =   19
            Top             =   210
            Width           =   600
         End
         Begin VB.Label lblImages 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Images"
            Height          =   195
            Left            =   1605
            TabIndex        =   18
            Top             =   1410
            Width           =   525
         End
         Begin VB.Label lblBackImage 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Back Image"
            Height          =   195
            Left            =   345
            TabIndex        =   17
            Top             =   2220
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frmStyleContainerStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private IsUpdating As Boolean
Private IsUpdatingData As Boolean
Private SelCIndex As Integer

Private Sub cmbBorder_Click()

    UpdateMenuItem

End Sub

Private Sub cmbFX_Click()

    UpdateMenuItem

End Sub

Private Sub cmdChange_Click()

    SelImage.FileName = picBackImage.Tag
    frmRscImages.Show vbModal
    
    With SelImage
        If .IsValid Then SetBackImage .FileName
    End With
    
    UpdateMenuItem

End Sub

Private Sub cmdColor_Click(Index As Integer)

    BuildUsedColorsArray

    With cmdColor(Index)
        SelColor = .Tag
        SelColor_CanBeTransparent = Not (Index = 1 Or Index = 3)
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdColor(Index)
    End With
    
    UpdateMenuItem

End Sub

Private Sub cmdRemove_Click()

    SetBackImage ""
    UpdateMenuItem

End Sub

Private Sub Form_Load()

    SetupCharset Me
    LocalizeUI

End Sub

Friend Sub UpdateUI(g As MenuGrp)

    If IsUpdatingData Then Exit Sub
    IsUpdating = True

    With g
        SetColor .bColor, cmdColor(0)
        SetColor .Corners.topCorner, cmdColor(5)
        SetColor .Corners.rightCorner, cmdColor(6)
        SetColor .Corners.bottomCorner, cmdColor(7)
        SetColor .Corners.leftCorner, cmdColor(8)
        
        cmbBorder.Text = .frameBorder
        cmbFX.ListIndex = .BorderStyle
        
        '-----------
        
        sbCImage(0).Tag = .CornersImages.gcTopLeft
        sbCImage(1).Tag = .CornersImages.gcTopCenter
        sbCImage(2).Tag = .CornersImages.gcTopRight
        sbCImage(3).Tag = .CornersImages.gcLeft
        sbCImage(4).Tag = .CornersImages.gcRight
        sbCImage(5).Tag = .CornersImages.gcBottomLeft
        sbCImage(6).Tag = .CornersImages.gcBottomCenter
        sbCImage(7).Tag = .CornersImages.gcBottomRight
        
        SetBackImage .Image
    End With
    
    sbCImage_Click 0
    
    IsUpdating = False

End Sub

Private Sub SetBackImage(fn As String)

    picBackImage.Tag = fn
    TileImage fn, picBackImage

End Sub

Private Sub LocalizeUI()

    cmbFX.Clear
    cmbFX.AddItem GetLocalizedStr(110)
    cmbFX.AddItem GetLocalizedStr(430)
    cmbFX.AddItem GetLocalizedStr(431)
    cmbFX.AddItem GetLocalizedStr(670)
    cmbFX.AddItem GetLocalizedStr(671)

    lblBorderSize.Caption = GetLocalizedStr(206)
    
    lblCorners.Caption = GetLocalizedStr(667)
    
    lblGBackColor.Caption = GetLocalizedStr(182)
    
    sbChangeCImage.Caption = GetLocalizedStr(189)
    sbRemoveCImage.Caption = GetLocalizedStr(201)
    
    lblImages.Caption = GetLocalizedStr(532)
    
    cmdRemove.Caption = GetLocalizedStr(201)
    
    lblBackImage.Caption = GetLocalizedStr(205)

End Sub

Friend Sub Form_Resize()

    On Error GoTo ExitSub

    svpMain.Move 0, 0, Width, Height - 3 * 15

    frameColor.Left = 8 * 15
    If Width < 472 * 15 Then
        frameColor.Width = Width - 18 * 15
        frameImage.Width = frameColor.Width
        frameImage.Left = frameColor.Left
        frameImage.Top = frameColor.Top + frameColor.Height + 10 * 15
    Else
        frameColor.Width = Width / 3 + 30 * 15
        frameImage.Left = frameColor.Left + frameColor.Width + 10 * 15
        frameImage.Width = Width - frameImage.Left - 10 * 15
        frameImage.Top = frameColor.Top
    End If
    
    uc3DLineColor.Width = frameColor.Width - 2 * uc3DLineColor.Left
    uc3DLineImage.Width = frameImage.Width - 2 * uc3DLineImage.Left
    
    svpMain.Refresh
    
ExitSub:

End Sub

Private Sub sbCImage_Click(Index As Integer)

    SelCIndex = Index

    Select Case Index
        Case 0
            lblSelCImage.Caption = GetLocalizedStr(716)
        Case 1
            lblSelCImage.Caption = GetLocalizedStr(717)
        Case 2
            lblSelCImage.Caption = GetLocalizedStr(718)
        Case 3
            lblSelCImage.Caption = GetLocalizedStr(719)
        Case 4
            lblSelCImage.Caption = GetLocalizedStr(720)
        Case 5
            lblSelCImage.Caption = GetLocalizedStr(721)
        Case 6
            lblSelCImage.Caption = GetLocalizedStr(722)
        Case 7
            lblSelCImage.Caption = GetLocalizedStr(723)
    End Select
    
    picCorner.Picture = LoadPictureRes(sbCImage(Index).Tag)
    
    UpdateCColors

End Sub

Private Sub UpdateCColors()

    Dim s As OLE_COLOR
    Dim i As Integer

    For i = 0 To sbCImage.Count - 1
        s = IIf(LenB(sbCImage(i).Tag) = 0, &HE0E0E0, &HFF0000)
        sbCImage(i).BackColor = IIf(i = SelCIndex, IIf(s = &HFF0000, &H800000, &H808080), s)
    Next i
    
    sbRemoveCImage.Enabled = (sbCImage(SelCIndex).BackColor = &H800000)
    
    UpdateMenuItem

End Sub

Private Sub UpdateMenuItem()

    If IsUpdating Then Exit Sub
    IsUpdatingData = True
    frmMain.UpdateItemData GetLocalizedStr(189) + cSep + "Container Style", False, True
    IsUpdatingData = False

End Sub

Private Sub sbRemoveCImage_Click()

    sbCImage(SelCIndex).Tag = vbNullString
    picCorner.Tag = vbNullString
    picCorner.Picture = LoadPicture()
    
    UpdateCColors

End Sub

Private Sub sbChangeCImage_Click()

    SelImage.FileName = sbCImage(SelCIndex).Tag
    frmRscImages.Show vbModal
    
    With SelImage
        If .IsValid Then
            sbCImage(SelCIndex).Tag = .FileName
            picCorner.Tag = .FileName
            picCorner.Picture = LoadPictureRes(picCorner.Tag)
        End If
    End With
    
    UpdateCColors

End Sub

