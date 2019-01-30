VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{C2F46FB4-62DF-499A-9E3D-EABC8CE04899}#72.0#0"; "SmartViewport.ocx"
Begin VB.Form frmStyleContainerSize 
   ClientHeight    =   3225
   ClientLeft      =   4830
   ClientTop       =   6900
   ClientWidth     =   8505
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
   ScaleHeight     =   3225
   ScaleWidth      =   8505
   Begin SmartViewportProject.SmartViewport svpMain 
      Height          =   2610
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8130
      _ExtentX        =   14340
      _ExtentY        =   4604
      ScrollBarType   =   1
      ScrollLeftRight =   0   'False
      ScrollSmallChange=   300
      ButtonChange    =   300
      Begin VB.Frame frameMargins 
         Caption         =   "Margins"
         Height          =   2370
         Left            =   0
         TabIndex        =   17
         Top             =   15
         Width           =   1740
         Begin VB.TextBox txtSpacing 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   24
            Text            =   "123"
            Top             =   1800
            Width           =   420
         End
         Begin VB.TextBox txtMarginH 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   19
            Text            =   "123"
            Top             =   510
            Width           =   420
         End
         Begin VB.TextBox txtMarginV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   18
            Text            =   "123"
            Top             =   1155
            Width           =   420
         End
         Begin MSComCtl2.UpDown UpDown2 
            Height          =   285
            Left            =   1140
            TabIndex        =   20
            Top             =   1155
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtMarginV"
            BuddyDispid     =   196612
            OrigLeft        =   2520
            OrigTop         =   630
            OrigRight       =   2715
            OrigBottom      =   885
            Max             =   20
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown1 
            Height          =   285
            Left            =   1140
            TabIndex        =   21
            Top             =   510
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtMarginH"
            BuddyDispid     =   196611
            OrigLeft        =   1260
            OrigTop         =   600
            OrigRight       =   1455
            OrigBottom      =   900
            Max             =   20
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin MSComCtl2.UpDown UpDown3 
            Height          =   285
            Left            =   1140
            TabIndex        =   25
            Top             =   1800
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtSpacing"
            BuddyDispid     =   196610
            OrigLeft        =   2520
            OrigTop         =   630
            OrigRight       =   2715
            OrigBottom      =   885
            Max             =   20
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   420
            Picture         =   "frmStyleContainerSize.frx":0000
            Top             =   1830
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spacing"
            Height          =   195
            Left            =   420
            TabIndex        =   26
            Top             =   1575
            Width           =   555
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horizontal"
            Height          =   195
            Left            =   420
            TabIndex        =   23
            Top             =   300
            Width           =   720
         End
         Begin VB.Label lblV 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vertical"
            Height          =   195
            Left            =   420
            TabIndex        =   22
            Top             =   930
            Width           =   525
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   420
            Picture         =   "frmStyleContainerSize.frx":014A
            Top             =   525
            Width           =   240
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   420
            Picture         =   "frmStyleContainerSize.frx":04D4
            Top             =   1170
            Width           =   240
         End
      End
      Begin VB.Frame frameSize 
         Caption         =   "Size"
         Height          =   1905
         Left            =   1920
         TabIndex        =   1
         Top             =   0
         Width           =   5940
         Begin VB.PictureBox picHeight 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   2805
            ScaleHeight     =   1335
            ScaleWidth      =   2475
            TabIndex        =   9
            Top             =   345
            Width           =   2475
            Begin VB.CommandButton cmdScrolling 
               Caption         =   "Scrolling..."
               Enabled         =   0   'False
               Height          =   285
               Left            =   1312
               TabIndex        =   14
               Top             =   1050
               Width           =   1170
            End
            Begin VB.TextBox txtHeight 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   862
               TabIndex        =   13
               Text            =   "0"
               Top             =   720
               Width           =   405
            End
            Begin VB.OptionButton opHeight 
               Caption         =   "Manual"
               Height          =   225
               Index           =   2
               Left            =   607
               TabIndex        =   12
               Top             =   480
               Width           =   1920
            End
            Begin VB.OptionButton opHeight 
               Caption         =   "Background Image"
               Height          =   225
               Index           =   1
               Left            =   607
               TabIndex        =   11
               Top             =   240
               Width           =   1920
            End
            Begin VB.OptionButton opHeight 
               Caption         =   "Auto"
               Height          =   225
               Index           =   0
               Left            =   607
               TabIndex        =   10
               Top             =   0
               Value           =   -1  'True
               Width           =   1920
            End
            Begin SmartButtonProject.SmartButton cmdAutoHeight 
               Height          =   300
               Left            =   1312
               TabIndex        =   15
               Top             =   720
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               Caption         =   "Calculate"
               Picture         =   "frmStyleContainerSize.frx":085E
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
               Enabled         =   0   'False
            End
            Begin VB.Label lblGSHeight 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Height"
               Height          =   195
               Left            =   0
               TabIndex        =   16
               Top             =   375
               Width           =   465
            End
            Begin VB.Image imgGSHeight 
               Appearance      =   0  'Flat
               Height          =   240
               Left            =   112
               Picture         =   "frmStyleContainerSize.frx":0BF8
               Top             =   150
               Width           =   240
            End
         End
         Begin VB.PictureBox picWidth 
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1035
            Left            =   165
            ScaleHeight     =   1035
            ScaleWidth      =   2700
            TabIndex        =   2
            Top             =   345
            Width           =   2700
            Begin VB.OptionButton opWidth 
               Caption         =   "Auto"
               Height          =   225
               Index           =   0
               Left            =   585
               TabIndex        =   6
               Top             =   0
               Width           =   1935
            End
            Begin VB.OptionButton opWidth 
               Caption         =   "Background Image"
               Height          =   225
               Index           =   1
               Left            =   585
               TabIndex        =   5
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtWidth 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   840
               TabIndex        =   4
               Text            =   "0"
               Top             =   735
               Width           =   405
            End
            Begin VB.OptionButton opWidth 
               Caption         =   "Manual"
               Height          =   225
               Index           =   2
               Left            =   585
               TabIndex        =   3
               Top             =   495
               Width           =   1935
            End
            Begin SmartButtonProject.SmartButton cmdAutoWidth 
               Height          =   300
               Left            =   1290
               TabIndex        =   7
               Top             =   735
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               Caption         =   "Calculate"
               Picture         =   "frmStyleContainerSize.frx":0F82
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
               Enabled         =   0   'False
            End
            Begin VB.Label lblGSWidth 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Width"
               Height          =   195
               Left            =   0
               TabIndex        =   8
               Top             =   450
               Width           =   420
            End
            Begin VB.Image imgGSWidth 
               Appearance      =   0  'Flat
               Height          =   240
               Left            =   90
               Picture         =   "frmStyleContainerSize.frx":131C
               Top             =   225
               Width           =   240
            End
         End
      End
   End
End
Attribute VB_Name = "frmStyleContainerSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim MenuItem As MenuGrp
Dim IsUpdating As Boolean

Private Sub Form_Load()

    SetupCharset Me
    LocalizeUI

End Sub

Friend Sub UpdateUI(g As MenuGrp)

    IsUpdating = True

    MenuItem = g
    With g
        txtMarginH.Text = .ContentsMarginH
        txtMarginV.Text = .ContentsMarginV
        
        If LenB(.Image) = 0 Then
            opWidth(1).Enabled = False
            opHeight(1).Enabled = False
        End If
        
        txtSpacing.Text = .Leading
        
        Select Case .fWidth
            Case -1
                opWidth(IIf(LenB(.Image) = 0, 0, 1)).Value = True
            Case 0
                opWidth(0).Value = True
            Case Else
                opWidth(2).Value = True
                txtWidth.Text = .fWidth
        End Select
        If .scrolling.MaxHeight > 0 Then .fHeight = .scrolling.MaxHeight
        Select Case .fHeight
            Case -1
                opHeight(IIf(LenB(.Image) = 0, 0, 1)).Value = True
            Case 0
                opHeight(0).Value = True
            Case Else
                opHeight(2).Value = True
                txtHeight.Text = .fHeight
        End Select
        
        cmdScrolling.Enabled = opHeight(2).Value And (g.fHeight < GetOptimalHeight) And (Val(txtHeight.Text) > 0)
    End With
    
    IsUpdating = False

End Sub

Friend Sub Form_Resize()

    On Error GoTo ExitSub

    svpMain.Move 0, 0, Width, Height - 3 * 15

    frameMargins.Left = 8 * 15
    If Width < 545 * 15 Then
        frameMargins.Width = Width - 18 * 15
        frameSize.Width = frameMargins.Width
        If Width < 390 * 15 Then
            picHeight.Move picWidth.Left, picWidth.Top + picWidth.Height + 15 * 15
            frameSize.Height = picHeight.Top + picHeight.Height + picWidth.Top
        Else
            frameSize.Height = 1905
            picHeight.Move picWidth.Left + picWidth.Width + 10 * 15, picWidth.Top
            frameSize.Left = frameMargins.Left
        End If
        frameSize.Top = frameMargins.Top + frameMargins.Height + 10 * 15
        frameSize.Left = frameMargins.Left
    Else
        frameMargins.Height = 2370
        frameSize.Height = frameMargins.Height
        frameMargins.Width = Width / 4 + 10 * 15
        frameSize.Left = frameMargins.Left + frameMargins.Width + 10 * 15
        frameSize.Width = Width - frameSize.Left - 10 * 15
        frameSize.Top = frameMargins.Top
    End If
    
    svpMain.Refresh
    
ExitSub:

End Sub

Private Sub txtMarginH_Change()

    UpdateMenuItem

End Sub

Private Sub UpdateMenuItem()

    If IsUpdating Then Exit Sub
    frmMain.UpdateItemData GetLocalizedStr(189) + cSep + "Container Size", True, True

End Sub

Private Sub txtMarginH_GotFocus()

    SelAll txtMarginH

End Sub

Private Sub txtMarginH_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtMarginV_Change()

    UpdateMenuItem

End Sub

Private Sub txtMarginV_GotFocus()

    SelAll txtMarginV

End Sub

Private Sub txtMarginV_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub LocalizeUI()

    frameMargins.Caption = GetLocalizedStr(209)
    
    lblH.Caption = GetLocalizedStr(211)
    lblV.Caption = GetLocalizedStr(210)
    
    opWidth(0).Caption = GetLocalizedStr(185)
    opWidth(1).Caption = GetLocalizedStr(199)
    opWidth(2).Caption = GetLocalizedStr(224)
    opHeight(0).Caption = GetLocalizedStr(185)
    opHeight(1).Caption = GetLocalizedStr(199)
    opHeight(2).Caption = GetLocalizedStr(224)
    lblGSWidth.Caption = GetLocalizedStr(428): lblGSWidth.Left = (imgGSWidth.Width - lblGSWidth.Width) / 2 + imgGSWidth.Left
    lblGSHeight.Caption = GetLocalizedStr(429): lblGSHeight.Left = (imgGSHeight.Width - lblGSHeight.Width) / 2 + imgGSHeight.Left
    cmdAutoWidth.Caption = GetLocalizedStr(225)
    cmdAutoHeight.Caption = GetLocalizedStr(225)
       
    FixContolsWidth Me
    
    cmdAutoWidth.Width = SetCtrlWidth(cmdAutoWidth)
    cmdAutoHeight.Width = SetCtrlWidth(cmdAutoHeight)
    
End Sub

Private Sub cmdAutoHeight_Click()

    txtHeight.Text = GetOptimalHeight
    txtHeight.SetFocus
    
    UpdateMenuItem

End Sub

Private Function GetOptimalHeight() As Integer

    Dim id As Integer
    
    id = GetIDByName(MenuItem.Name)
    MenuGrps(id).fHeight = 0
    GetOptimalHeight = GetDivHeight(id)
    MenuGrps(id).fHeight = MenuItem.fHeight

End Function

Private Sub cmdAutoWidth_Click()

    Dim id As Integer
    
    id = GetIDByName(MenuItem.Name)

    MenuGrps(id).fWidth = 0
    txtWidth.Text = GetDivWidth(id)
    txtWidth.SetFocus
    MenuGrps(id).fWidth = MenuItem.fWidth
    
    UpdateMenuItem

End Sub

Private Sub cmdScrolling_Click()

    On Error Resume Next
   
    frmGrpScrolling.Show vbModal
    txtHeight.SetFocus
    
    UpdateMenuItem

End Sub

Private Sub opHeight_Click(Index As Integer)

    txtHeight.Enabled = opHeight(2).Value
    cmdAutoHeight.Enabled = opHeight(2).Value
    
    If IsUpdating Then Exit Sub
    If Index = 2 And Val(txtHeight.Text) = 0 Then
        IsUpdating = True
        cmdAutoHeight_Click
        IsUpdating = False
    End If
    
    UpdateMenuItem

End Sub

Private Sub opWidth_Click(Index As Integer)

    txtWidth.Enabled = opWidth(2).Value
    cmdAutoWidth.Enabled = opWidth(2).Value
    
    If IsUpdating Then Exit Sub
    If Index = 2 And Val(txtWidth.Text) = 0 Then
        IsUpdating = True
        cmdAutoWidth_Click
        IsUpdating = False
    End If
    
    UpdateMenuItem

End Sub

Private Sub txtHeight_Change()

    UpdateMenuItem

End Sub

Private Sub txtHeight_GotFocus()

    SelAll txtHeight

End Sub

Private Sub txtHeight_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub

Private Sub txtSpacing_Change()

    UpdateMenuItem

End Sub

Private Sub txtSpacing_GotFocus()

    SelAll txtSpacing

End Sub

Private Sub txtSpacing_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
    
End Sub

Private Sub txtWidth_Change()

    UpdateMenuItem

End Sub

Private Sub txtWidth_GotFocus()

    SelAll txtWidth

End Sub

Private Sub txtWidth_KeyPress(KeyAscii As Integer)

    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 Then
        KeyAscii = 0
    End If

End Sub
