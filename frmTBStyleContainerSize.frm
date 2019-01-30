VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Object = "{C2F46FB4-62DF-499A-9E3D-EABC8CE04899}#72.0#0"; "SmartViewport.ocx"
Begin VB.Form frmTBStyleContainerSize 
   ClientHeight    =   3165
   ClientLeft      =   3960
   ClientTop       =   5865
   ClientWidth     =   8310
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
   ScaleHeight     =   3165
   ScaleWidth      =   8310
   Begin SmartViewportProject.SmartViewport svpMain 
      Height          =   2655
      Left            =   90
      TabIndex        =   0
      Top             =   105
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   4683
      ScrollBarType   =   1
      ScrollLeftRight =   0   'False
      ScrollSmallChange=   300
      ButtonChange    =   300
      Begin VB.Frame frameMargins 
         Caption         =   "Margins"
         Height          =   2370
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   1740
         Begin VB.TextBox txtMarginV 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   19
            Text            =   "123"
            Top             =   1155
            Width           =   420
         End
         Begin VB.TextBox txtMarginH 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   18
            Text            =   "123"
            Top             =   510
            Width           =   420
         End
         Begin VB.TextBox txtSpacing 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   720
            TabIndex        =   17
            Text            =   "123"
            Top             =   1800
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
            BuddyDispid     =   196610
            OrigLeft        =   1140
            OrigTop         =   1155
            OrigRight       =   1380
            OrigBottom      =   1440
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
            TabIndex        =   22
            Top             =   1800
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   503
            _Version        =   393216
            BuddyControl    =   "txtSpacing"
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
         Begin VB.Image Image2 
            Height          =   240
            Left            =   420
            Picture         =   "frmTBStyleContainerSize.frx":0000
            Top             =   1170
            Width           =   240
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   420
            Picture         =   "frmTBStyleContainerSize.frx":038A
            Top             =   525
            Width           =   240
         End
         Begin VB.Label lblV 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Vertical"
            Height          =   195
            Left            =   420
            TabIndex        =   25
            Top             =   930
            Width           =   525
         End
         Begin VB.Label lblH 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Horizontal"
            Height          =   195
            Left            =   420
            TabIndex        =   24
            Top             =   300
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spacing"
            Height          =   195
            Left            =   420
            TabIndex        =   23
            Top             =   1575
            Width           =   795
         End
         Begin VB.Image Image3 
            Height          =   240
            Left            =   420
            Picture         =   "frmTBStyleContainerSize.frx":0714
            Top             =   1830
            Width           =   240
         End
      End
      Begin VB.Frame frameSize 
         Caption         =   "Size"
         Height          =   2370
         Left            =   1920
         TabIndex        =   1
         Top             =   0
         Width           =   6180
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
            TabIndex        =   10
            Top             =   345
            Width           =   2475
            Begin VB.TextBox txtHeight 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   862
               TabIndex        =   13
               Text            =   "0"
               Top             =   480
               Width           =   405
            End
            Begin VB.OptionButton opHeight 
               Caption         =   "Manual"
               Height          =   225
               Index           =   1
               Left            =   607
               TabIndex        =   12
               Top             =   240
               Width           =   1920
            End
            Begin VB.OptionButton opHeight 
               Caption         =   "Auto"
               Height          =   225
               Index           =   0
               Left            =   607
               TabIndex        =   11
               Top             =   0
               Value           =   -1  'True
               Width           =   1920
            End
            Begin SmartButtonProject.SmartButton cmdAutoHeight 
               Height          =   300
               Left            =   1305
               TabIndex        =   14
               Top             =   480
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               Caption         =   "Calculate"
               Picture         =   "frmTBStyleContainerSize.frx":085E
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
               TabIndex        =   15
               Top             =   375
               Width           =   465
            End
            Begin VB.Image imgGSHeight 
               Appearance      =   0  'Flat
               Height          =   240
               Left            =   112
               Picture         =   "frmTBStyleContainerSize.frx":0BF8
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
            TabIndex        =   3
            Top             =   345
            Width           =   2700
            Begin VB.OptionButton opWidth 
               Caption         =   "Auto"
               Height          =   225
               Index           =   0
               Left            =   585
               TabIndex        =   7
               Top             =   0
               Width           =   1935
            End
            Begin VB.OptionButton opWidth 
               Caption         =   "Match Group's Width"
               Height          =   225
               Index           =   1
               Left            =   585
               TabIndex        =   6
               Top             =   240
               Width           =   1935
            End
            Begin VB.TextBox txtWidth 
               Alignment       =   1  'Right Justify
               Enabled         =   0   'False
               Height          =   285
               Left            =   840
               TabIndex        =   5
               Text            =   "0"
               Top             =   735
               Width           =   405
            End
            Begin VB.OptionButton opWidth 
               Caption         =   "Manual"
               Height          =   225
               Index           =   2
               Left            =   585
               TabIndex        =   4
               Top             =   495
               Width           =   1935
            End
            Begin SmartButtonProject.SmartButton cmdAutoWidth 
               Height          =   300
               Left            =   1290
               TabIndex        =   8
               Top             =   735
               Width           =   1170
               _ExtentX        =   2064
               _ExtentY        =   529
               Caption         =   "Calculate"
               Picture         =   "frmTBStyleContainerSize.frx":0F82
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
               TabIndex        =   9
               Top             =   450
               Width           =   420
            End
            Begin VB.Image imgGSWidth 
               Appearance      =   0  'Flat
               Height          =   240
               Left            =   90
               Picture         =   "frmTBStyleContainerSize.frx":131C
               Top             =   225
               Width           =   240
            End
         End
         Begin VB.CheckBox chkJustify 
            Caption         =   "Make all Items have the same width/height"
            Height          =   240
            Left            =   165
            TabIndex        =   2
            Top             =   1860
            Width           =   3510
         End
      End
   End
End
Attribute VB_Name = "frmTBStyleContainerSize"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsUpdating As Boolean

Private Sub chkJustify_Click()

    UpdateToolbar

End Sub

Private Sub cmdAutoHeight_Click()

    Dim r() As Integer

    r = GetTBHeight(ToolbarIndexByKey(frmMain.tvMapView.SelectedItem.key), True)
    txtHeight.Text = r(1)

End Sub

Private Sub cmdAutoWidth_Click()

    Dim r() As Integer

    r = GetTBWidth(ToolbarIndexByKey(frmMain.tvMapView.SelectedItem.key), True)
    txtWidth.Text = r(1)

End Sub

Private Sub Form_Load()

    SetupCharset Me
    'LocalizeUI
    FixCtrls4Skin Me

End Sub

Private Sub UpdateToolbar()

    If IsUpdating Then Exit Sub
    frmMain.UpdateToolbarData

End Sub

Friend Sub UpdateUI(tb As ToolbarDef)

    IsUpdating = True

    With tb
        txtMarginH.Text = .ContentsMarginH
        txtMarginV.Text = .ContentsMarginV
        txtSpacing.Text = .Separation
        
        Select Case .Width
            Case 0
                opWidth(0).Value = True
            Case 1
                opWidth(1).Value = True
            Case Else
                opWidth(2).Value = True
                txtWidth.Text = Abs(.Width)
        End Select
        Select Case .Height
            Case 0
                opHeight(0).Value = True
            Case Else
                opHeight(1).Value = True
                txtHeight.Text = Abs(.Height)
        End Select
        
        chkJustify.Value = IIf(.JustifyHotSpots, vbChecked, vbUnchecked)
    End With
    
    IsUpdating = False

End Sub

Friend Sub Form_Resize()

    On Error GoTo ExitSub

    svpMain.Move 0, 0, Width, Height - 3 * 15

    frameMargins.Left = 8 * 15
    If Width < 8010 Then
        frameMargins.Width = Width - 18 * 15
        frameSize.Width = frameMargins.Width
        If Width < 395 * 15 Then
            picHeight.Move picWidth.Left, picWidth.Top + picWidth.Height + 15 * 15
            frameSize.Height = picHeight.Top + picHeight.Height + picWidth.Top + 5 * 15
        Else
            frameSize.Height = 1905 + 12 * 15
            picHeight.Move picWidth.Left + picWidth.Width + 10 * 15, picWidth.Top
        End If
        frameSize.Left = frameMargins.Left
        frameSize.Top = frameMargins.Top + frameMargins.Height + 10 * 15
    Else
        frameMargins.Height = 2370
        frameSize.Height = frameMargins.Height
        frameMargins.Width = Width / 4 + 10 * 15
        frameSize.Left = frameMargins.Left + frameMargins.Width + 10 * 15
        frameSize.Width = Width - frameSize.Left - 10 * 15
        frameSize.Top = frameMargins.Top
    End If
    
    chkJustify.Top = picHeight.Top + picHeight.Height + 5 * 15
    
    svpMain.Refresh
    
ExitSub:

End Sub

Private Sub opHeight_Click(Index As Integer)

    txtHeight.Enabled = (Index = 1)
    cmdAutoHeight.Enabled = (Index = 1)
    
    If Index = 1 And Val(txtHeight.Text) = 0 Then
        IsUpdating = True
        cmdAutoHeight_Click
        IsUpdating = False
    End If

    UpdateToolbar

End Sub

Private Sub opWidth_Click(Index As Integer)

    txtWidth.Enabled = (Index = 2)
    cmdAutoWidth.Enabled = (Index = 2)
    
    If Index = 2 And Val(txtWidth.Text) = 0 Then
        IsUpdating = True
        cmdAutoWidth_Click
        IsUpdating = False
    End If

    UpdateToolbar

End Sub

Private Sub txtHeight_Change()

    UpdateToolbar

End Sub

Private Sub txtHeight_GotFocus()

    SelAll txtHeight

End Sub

Private Sub txtMarginH_Change()

    UpdateToolbar

End Sub

Private Sub txtMarginH_GotFocus()

    SelAll txtMarginH

End Sub

Private Sub txtMarginV_Change()

    UpdateToolbar

End Sub

Private Sub txtMarginV_GotFocus()

    SelAll txtMarginV

End Sub

Private Sub txtSpacing_Change()

    UpdateToolbar

End Sub

Private Sub txtSpacing_GotFocus()

    SelAll txtSpacing

End Sub

Private Sub txtWidth_Change()

    UpdateToolbar

End Sub

Private Sub txtWidth_GotFocus()

    SelAll txtWidth

End Sub
