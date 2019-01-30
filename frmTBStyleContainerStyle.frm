VERSION 5.00
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{C2F46FB4-62DF-499A-9E3D-EABC8CE04899}#72.0#0"; "SmartViewport.ocx"
Begin VB.Form frmTBStyleContainerStyle 
   ClientHeight    =   2505
   ClientLeft      =   2565
   ClientTop       =   5460
   ClientWidth     =   9975
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
   ScaleHeight     =   2505
   ScaleWidth      =   9975
   Begin SmartViewportProject.SmartViewport svpMain 
      Height          =   2100
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   9600
      _ExtentX        =   16933
      _ExtentY        =   3704
      ScrollBarType   =   1
      ScrollLeftRight =   0   'False
      ScrollSmallChange=   300
      ButtonChange    =   300
      Begin VB.Frame frameGeneral 
         Caption         =   "General"
         Height          =   1575
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3060
         Begin VB.ComboBox cmbStyle 
            Height          =   315
            ItemData        =   "frmTBStyleContainerStyle.frx":0000
            Left            =   1170
            List            =   "frmTBStyleContainerStyle.frx":000A
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   405
            Width           =   1665
         End
         Begin VB.Label lblTBStyle 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Toolbar Style"
            Height          =   195
            Left            =   150
            TabIndex        =   16
            Top             =   465
            Width           =   945
         End
      End
      Begin VB.Frame frameImage 
         Caption         =   "Image"
         Height          =   1575
         Left            =   6285
         TabIndex        =   9
         Top             =   0
         Width           =   2865
         Begin VB.PictureBox picTBImage 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00E0E0E0&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   2175
            ScaleHeight     =   450
            ScaleWidth      =   450
            TabIndex        =   10
            Top             =   315
            Width           =   480
         End
         Begin SmartButtonProject.SmartButton cmdChange 
            Height          =   240
            Left            =   1065
            TabIndex        =   11
            Top             =   315
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   423
            Caption         =   "Change"
            Picture         =   "frmTBStyleContainerStyle.frx":0024
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
            Left            =   1065
            TabIndex        =   12
            Top             =   555
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   423
            Caption         =   "Remove"
            Picture         =   "frmTBStyleContainerStyle.frx":03BE
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
         Begin VB.Label lblTBBackImage 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Back Image"
            Height          =   195
            Left            =   165
            TabIndex        =   13
            Top             =   435
            Width           =   825
         End
      End
      Begin VB.Frame frameColor 
         Caption         =   "Color"
         Height          =   1575
         Left            =   3210
         TabIndex        =   1
         Top             =   75
         Width           =   2910
         Begin VB.ComboBox cmbBorder 
            Height          =   315
            ItemData        =   "frmTBStyleContainerStyle.frx":0758
            Left            =   1095
            List            =   "frmTBStyleContainerStyle.frx":075A
            TabIndex        =   3
            Text            =   "cmbBorder"
            Top             =   660
            Width           =   1305
         End
         Begin VB.ComboBox cmbFX 
            Height          =   315
            ItemData        =   "frmTBStyleContainerStyle.frx":075C
            Left            =   1095
            List            =   "frmTBStyleContainerStyle.frx":0769
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1095
            Width           =   1305
         End
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   240
            Index           =   1
            Left            =   2460
            TabIndex        =   4
            Top             =   690
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
            ShowFocus       =   -1  'True
         End
         Begin SmartButtonProject.SmartButton cmdColor 
            Height          =   240
            Index           =   0
            Left            =   1095
            TabIndex        =   5
            Top             =   330
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
            ShowFocus       =   -1  'True
         End
         Begin VB.Label lblTBBorder 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Border"
            Height          =   195
            Left            =   540
            TabIndex        =   8
            Top             =   720
            Width           =   480
         End
         Begin VB.Label lblBorderStyle 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Border Style"
            Height          =   195
            Left            =   135
            TabIndex        =   7
            Top             =   1155
            Width           =   885
         End
         Begin VB.Label lblTBBackColor 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Back Color"
            Height          =   195
            Left            =   270
            TabIndex        =   6
            Top             =   345
            Width           =   750
         End
      End
   End
End
Attribute VB_Name = "frmTBStyleContainerStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsUpdating As Boolean

Private Sub UpdateToolbar()

    If IsUpdating Then Exit Sub
    frmMain.UpdateToolbarData

End Sub

Private Sub cmbBorder_Change()

    UpdateToolbar

End Sub

Private Sub cmbBorder_Click()

    UpdateToolbar

End Sub

Private Sub cmbFX_Click()

    UpdateToolbar

End Sub

Private Sub cmbStyle_Click()

    UpdateToolbar

End Sub

Private Sub cmdTBBackColor_Click()

End Sub

Private Sub cmdChange_Click()

    SelImage.FileName = picTBImage.Tag
    SelImage.SupportsFlash = True
    frmRscImages.Show vbModal
    With SelImage
        If .IsValid Then SetTBPicture .FileName
    End With
    
    UpdateToolbar
    Me.SetFocus

End Sub

Private Sub SetTBPicture(fn As String)

    On Error Resume Next

    picTBImage.Tag = fn
    TileImage fn, picTBImage

End Sub

Private Sub cmdColor_Click(Index As Integer)

    BuildUsedColorsArray

    With cmdColor(Index)
        SelColor = .Tag
        SelColor_CanBeTransparent = (Index = 0)
        frmColorPicker.Show vbModal
        SetColor SelColor, cmdColor(Index)
    End With
    
    UpdateToolbar
    
    Me.SetFocus

End Sub

Private Sub cmdRemove_Click()

    SetTBPicture ""
    UpdateToolbar

End Sub

Private Sub Form_Load()

    Dim i As Integer

    SetupCharset Me
    LocalizeUI
    
    cmbBorder.AddItem GetLocalizedStr(110)
    For i = 1 To 10
        cmbBorder.AddItem CStr(i)
    Next i

End Sub

Friend Sub UpdateUI(tb As ToolbarDef)

    IsUpdating = True

    With tb
        cmbStyle.ListIndex = .Style

        SetColor .BackColor, cmdColor(0)
        cmbBorder.Text = .border
        cmbFX.ListIndex = .BorderStyle
        SetColor .BorderColor, cmdColor(1)
        
        SetTBPicture .Image
    End With
    
    IsUpdating = False

End Sub

Private Sub Form_Resize()

    Dim minW As Single
    Dim c As Control
    Dim w As Single
    
    On Error GoTo ExitSub
    
    svpMain.Move 0, 0, Width, Height - 3 * 15
    w = Width - 8 * 15
    
    minW = 185 * 15 + 2 * frameGeneral.Left + 5 * 15
    
    frameGeneral.Left = 8 * 15
    If w >= minW * 3 Then
        frameGeneral.Height = frameColor.Height
        frameGeneral.Width = w / 3 - 10 * 15
        
        frameColor.Top = frameGeneral.Top
        frameColor.Width = frameGeneral.Width
        frameColor.Left = frameGeneral.Left + frameGeneral.Width + 10 * 15
        
        frameImage.Height = frameColor.Height
        frameImage.Top = frameGeneral.Top
        frameImage.Width = frameGeneral.Width
        frameImage.Left = frameColor.Left + frameColor.Width + 10 * 15
    End If
    
    If w >= minW * 2 And w < minW * 3 Then
        frameGeneral.Height = frameColor.Height
        frameGeneral.Width = w / 2 - 10 * 15
        
        frameColor.Top = frameGeneral.Top
        frameColor.Width = frameGeneral.Width
        frameColor.Left = frameGeneral.Left + frameGeneral.Width + 10 * 15
        
        frameImage.Height = 75 * 15
        frameImage.Width = w - frameGeneral.Left * 2 + 6 * 15
        frameImage.Left = frameGeneral.Left
        frameImage.Top = frameGeneral.Top + frameGeneral.Height + 10 * 15
    End If
    
    If w <= minW * 2 Then
        frameGeneral.Height = 75 * 15
        frameGeneral.Width = w - frameGeneral.Left * 2 + 6 * 15
        
        frameColor.Left = frameGeneral.Left
        frameColor.Top = frameGeneral.Top + frameGeneral.Height + 10 * 15
        frameColor.Width = frameGeneral.Width
        
        frameImage.Height = 75 * 15
        frameImage.Width = frameGeneral.Width
        frameImage.Left = frameGeneral.Left
        frameImage.Top = frameColor.Top + frameColor.Height + 10 * 15
    End If
    
    svpMain.Refresh
    
ExitSub:

End Sub

Private Sub LocalizeUI()

    cmbFX.Clear
    cmbFX.AddItem GetLocalizedStr(110)
    cmbFX.AddItem GetLocalizedStr(430)
    cmbFX.AddItem GetLocalizedStr(431)
    cmbFX.AddItem GetLocalizedStr(670)
    cmbFX.AddItem GetLocalizedStr(671)

End Sub

