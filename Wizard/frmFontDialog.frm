VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmFontDialog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Font"
   ClientHeight    =   4590
   ClientLeft      =   5970
   ClientTop       =   3855
   ClientWidth     =   3615
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFontDialog.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   3615
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
      Left            =   180
      TabIndex        =   8
      Top             =   2055
      Width           =   1980
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
      Left            =   180
      TabIndex        =   7
      Top             =   2280
      Width           =   1980
   End
   Begin VB.ComboBox cmbSize 
      Height          =   315
      Left            =   2250
      TabIndex        =   6
      Text            =   "cmbSize"
      Top             =   2055
      Width           =   690
   End
   Begin VB.PictureBox picSample 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000009&
      Height          =   1005
      Left            =   30
      ScaleHeight     =   945
      ScaleWidth      =   3465
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2850
      Width           =   3525
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   1560
      TabIndex        =   3
      Top             =   4155
      Width           =   900
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2655
      TabIndex        =   2
      Top             =   4155
      Width           =   900
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
      Left            =   180
      TabIndex        =   1
      Top             =   2505
      Width           =   1980
   End
   Begin VB.ListBox lstFont 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1425
      ItemData        =   "frmFontDialog.frx":014A
      Left            =   30
      List            =   "frmFontDialog.frx":014C
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   270
      Width           =   3525
   End
   Begin xfxLine3D.ucLine3D uc3DLine3 
      Height          =   30
      Left            =   0
      TabIndex        =   4
      Top             =   4005
      Width           =   3585
      _ExtentX        =   6324
      _ExtentY        =   53
   End
   Begin VB.Label lblPoints 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "(0 points)"
      ForeColor       =   &H8000000C&
      Height          =   195
      Left            =   2250
      TabIndex        =   14
      Top             =   2415
      Width           =   690
   End
   Begin VB.Label lblName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   30
      TabIndex        =   13
      Top             =   45
      Width           =   405
   End
   Begin VB.Label lblStyle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Style"
      Height          =   195
      Left            =   30
      TabIndex        =   12
      Top             =   1830
      Width           =   360
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size"
      Height          =   195
      Left            =   2145
      TabIndex        =   11
      Top             =   1830
      Width           =   285
   End
   Begin VB.Label lblSample 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Sample"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   4275
      Visible         =   0   'False
      Width           =   510
   End
   Begin VB.Label lblUnit 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "pixels"
      Height          =   195
      Left            =   3000
      TabIndex        =   9
      Top             =   2100
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

    SelFont.Size = Val(cmbSize.Text)
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

Private Sub Form_Load()

    Dim i As Integer

    CenterForm Me
    
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
    Dim s1 As Long
    
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

    Dim Caption As String

    If IsUpdating Or cmbSize.Text = "" Then Exit Sub

    With lblSample.Font
        .Name = lstFont.Text
        .Italic = (chkItalic.Value = vbChecked)
        .Bold = (chkBold.Value = vbChecked)
        .Underline = (chkUnderline.Value = vbChecked)
        .Size = px2pt(SelFont.Size)
    End With
    
    If Caption = "" Then
        lblSample.Caption = lblSample.FontName
    Else
        lblSample.Caption = Caption
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
    picSample.Print lblSample.Caption
    
    lblPoints.Caption = "(" & px2pt(SelFont.Size) & " points)"

End Sub
