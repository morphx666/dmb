VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmFontDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Font"
   ClientHeight    =   4815
   ClientLeft      =   7605
   ClientTop       =   5340
   ClientWidth     =   3705
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
   Icon            =   "frmFontDialog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   3705
   ShowInTaskbar   =   0   'False
   Begin VB.ListBox lstFont 
      Height          =   1530
      ItemData        =   "frmFontDialog.frx":014A
      Left            =   90
      List            =   "frmFontDialog.frx":014C
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   285
      Width           =   3525
   End
   Begin VB.CheckBox chkUnderline 
      Caption         =   "Underline"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   9
      Top             =   2700
      Width           =   1980
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
      Left            =   2715
      TabIndex        =   14
      Top             =   4350
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
      Left            =   1680
      TabIndex        =   13
      Top             =   4350
      Width           =   900
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   60
      TabIndex        =   11
      Top             =   4200
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   53
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      Height          =   1005
      Left            =   90
      ScaleHeight     =   945
      ScaleWidth      =   3465
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   3045
      Width           =   3525
   End
   Begin VB.ComboBox cmbSize 
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
      Left            =   2310
      TabIndex        =   5
      Text            =   "cmbSize"
      Top             =   2250
      Width           =   690
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Bold"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   2475
      Width           =   1980
   End
   Begin VB.CheckBox chkItalic 
      Caption         =   "Italic"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   240
      TabIndex        =   4
      Top             =   2250
      Width           =   1980
   End
   Begin VB.Label lblPoints 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0 points)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000C&
      Height          =   195
      Left            =   2310
      TabIndex        =   8
      Top             =   2595
      Width           =   690
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
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
      Left            =   3060
      TabIndex        =   6
      Top             =   2295
      Width           =   405
   End
   Begin VB.Label lblSample 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sample"
      Height          =   210
      Left            =   240
      TabIndex        =   12
      Top             =   4470
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   210
      Left            =   2310
      TabIndex        =   3
      Top             =   2010
      Width           =   315
   End
   Begin VB.Label lblStyle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style"
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   2010
      Width           =   360
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   60
      Width           =   405
   End
End
Attribute VB_Name = "frmFontDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsUpdating As Boolean

Private Sub chkBold_Click()

    UpdateSample

End Sub

Private Sub chkItalic_Click()

    UpdateSample

End Sub

Private Sub chkUnderline_Click()

    UpdateSample

End Sub

Private Sub cmbSize_Change()

    cmbSize_Click

End Sub

Private Sub cmbSize_Click()

    Dim newVal As Integer
    
    On Error Resume Next
    
    newVal = Val(cmbSize.Text)
    If newVal <= 0 Then newVal = 1

    SelFont.Size = newVal
    UpdateSample

End Sub

Private Sub cmdCancel_Click()

    SelFont.IsValid = False
    SelFont.IsSubst = False
    Hide

End Sub

Private Sub cmdOK_Click()

    With SelFont
        .Name = lblSample.FontName
        .Italic = lblSample.FontItalic
        .Bold = lblSample.FontBold
        .Size = Val(cmbSize.Text)
        .Underline = lblSample.FontUnderline
        .IsValid = True
        .IsSubst = False
    End With
    Hide

End Sub

Private Sub Form_Activate()

    Dim i As Integer

    For i = 0 To lstFont.ListCount - 1
        If lstFont.List(i) = SelFont.Name Then
            lstFont.Selected(i) = True
            Exit For
        End If
    Next i
    If lstFont.ListIndex = -1 Then lstFont.ListIndex = 0
    
    SelCurSize
    
    chkItalic.Value = Abs(SelFont.Italic)
    chkBold.Value = Abs(SelFont.Bold)
    chkUnderline.Value = Abs(SelFont.Underline)
    
    chkBold.Enabled = Not SelFont.IsSubst
    chkItalic.Enabled = Not SelFont.IsSubst
    chkUnderline.Enabled = Not SelFont.IsSubst
    cmbSize.Enabled = Not SelFont.IsSubst

    UpdateSample

End Sub

Private Sub SelCurSize()

    Dim i As Integer
    
ReStart:
    For i = 0 To cmbSize.ListCount - 1
        If SelFont.Size = Val(cmbSize.List(i)) Then
            cmbSize.ListIndex = i
            Exit Sub
        End If
    Next i
    
    For i = 0 To cmbSize.ListCount - 1
        If Val(cmbSize.List(i)) > SelFont.Size Then
            cmbSize.AddItem SelFont.Size, i
            GoTo ReStart
        End If
    Next i
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then showHelp "dialogs/font_selection.htm"

End Sub

Private Sub Form_Load()

    CenterForm Me
    LocalizeUI
    
    IsUpdating = True
    
    GetFontList
    
    IsUpdating = False

End Sub

Private Sub GetFontList()

    Dim hDc As Long
    
    lstFont.Clear
    
    'get the handle to the device context of the list to fill
    hDc = GetDC(lstFont.hwnd)
    
    'Add the fonts using the API and callback by calling
    'the EnumFontFamilies API, passing the AddressOf the
    'application-defined callback procedure EnumFontFamProc
    'and the list to fill
    EnumFontFamilies hDc, vbNullString, AddressOf EnumFontFamProc, lstFont
    
    'free the device context handle
    ReleaseDC lstFont.hwnd, hDc
    
End Sub

Private Sub lstFont_Click()

    GetValidFontSizes
    UpdateSample

End Sub

Private Sub GetValidFontSizes()

    Dim i As Integer
    
    On Error Resume Next
    
    With lblSample.Font
        .Name = lstFont.Text
        .Italic = (chkItalic.Value = vbChecked)
        .Bold = (chkBold.Value = vbChecked)
        .Underline = (chkUnderline.Value = vbChecked)

        cmbSize.Clear
        For i = 1 To 72
            .Size = px2pt(i)
            If pt2px(.Size) = i Then cmbSize.AddItem i
        Next i
    End With
    
    SelCurSize
    
End Sub

Private Sub UpdateSample()

    Dim caption As String

    If IsUpdating Or LenB(cmbSize.Text) = 0 Then Exit Sub

    With lblSample.Font
        .Name = lstFont.Text
        .Italic = (chkItalic.Value = vbChecked)
        .Bold = (chkBold.Value = vbChecked)
        .Underline = (chkUnderline.Value = vbChecked)
        .Size = px2pt(SelFont.Size)
    End With
    
    If IsGroup(frmMain.tvMenus.SelectedItem.key) Then caption = MenuGrps(GetID).caption
    If IsCommand(frmMain.tvMenus.SelectedItem.key) Then caption = MenuCmds(GetID).caption
    If LenB(caption) = 0 Then
        lblSample.caption = lblSample.FontName
    Else
        lblSample.caption = caption
    End If
    
    DoEvents
    
    With picSample
        .Cls
        .CurrentX = .Width / 2 - lblSample.Width / 2
        .CurrentY = .Height / 2 - lblSample.Height / 2
        .Font.Name = lblSample.Font.Name
        .Font.Italic = lblSample.Font.Italic
        .Font.Bold = lblSample.Font.Bold
        .Font.Size = lblSample.Font.Size
        .Font.Underline = lblSample.Font.Underline
    End With
    picSample.Print lblSample.caption
    
    lblPoints.caption = "(" & px2pt(SelFont.Size) & " points)"

End Sub

Private Sub LocalizeUI()

    lblName.caption = GetLocalizedStr(409)
    lblStyle.caption = GetLocalizedStr(613)
    lblSize.caption = GetLocalizedStr(203)
    lblUnit.caption = GetLocalizedStr(614)
    
    chkBold.caption = GetLocalizedStr(510)
    chkItalic.caption = GetLocalizedStr(511)
    chkUnderline.caption = GetLocalizedStr(512)

    cmdOK.caption = GetLocalizedStr(186)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdOK.Width = SetCtrlWidth(cmdOK)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If
    
    FixContolsWidth Me

End Sub
