VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmLivePreview 
   Caption         =   "Form1"
   ClientHeight    =   6390
   ClientLeft      =   3360
   ClientTop       =   5400
   ClientWidth     =   10305
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
   ScaleHeight     =   6390
   ScaleWidth      =   10305
   Begin VB.Frame frameControls 
      Height          =   1050
      Left            =   270
      TabIndex        =   1
      Top             =   2355
      Width           =   6240
      Begin VB.CheckBox chkUseTarget 
         Caption         =   "Use Target Document"
         Height          =   210
         Left            =   105
         TabIndex        =   3
         Top             =   675
         Width           =   2220
      End
      Begin VB.ComboBox cmbCharset 
         Height          =   315
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   225
         Width           =   2565
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Left            =   1440
         TabIndex        =   4
         Top             =   255
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
      Begin VB.Label lblBackColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Background Color"
         Height          =   195
         Left            =   105
         TabIndex        =   6
         Top             =   285
         Width           =   1260
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   1920
         X2              =   1920
         Y1              =   225
         Y2              =   540
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   1935
         X2              =   1935
         Y1              =   225
         Y2              =   540
      End
      Begin VB.Label lblEncoding 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encoding"
         Height          =   195
         Left            =   2100
         TabIndex        =   5
         Top             =   285
         Width           =   645
      End
   End
   Begin SHDocVwCtl.WebBrowser wbPreview 
      Height          =   1950
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1905
      ExtentX         =   3360
      ExtentY         =   3440
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmLivePreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdColor_Click()

    BuildUsedColorsArray

    SelColor = GetSetting(App.EXEName, "PreviewWinPos", "BackColor", vbWhite)
    SelColor_CanBeTransparent = False
    frmColorPicker.Show vbModal, Me

    If SelColor <> -1 Then
        SetColor SelColor, cmdColor
        SetDocBackColor

        SaveSetting App.EXEName, "PreviewWinPos", "BackColor", SelColor
    End If

End Sub

Private Sub Form_Load()

    SetColor GetSetting(App.EXEName, "PreviewWinPos", "BackColor", vbWhite), cmdColor
    
    SetDocBackColor

    FillCodepageCombo
    FixCtrls4Skin Me
    
End Sub

Private Sub FillCodepageCombo()

    Dim i As Integer

    cs = GetSysCharsets
    
    For i = 1 To UBound(cs)
        cmbCharset.AddItem cs(i).Description
        If cs(i).CodePage = Preferences.CodePage Then
            cmbCharset.ListIndex = cmbCharset.NewIndex
        End If
    Next i
    ResizeComboList cmbCharset

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If WindowState <> vbMinimized Then
        If Width > 2000 And Height > 2000 Then
            With wbPreview
                .Move 60, 60, Width - 235, Height - 330 - frameControls.Height
                frameControls.Move 60, .Top + .Height + 30, .Width, frameControls.Height
            End With
            With frameControls
                If IsSkinned Then Controls("dynpic1").Width = .Width - 60
            End With
        End If
        With cmbCharset
            If Width > .Left + 2565 + 300 Then
                .Width = 2565
            Else
                .Width = Width - .Left - 375
            End If
        End With
    End If

End Sub

Private Sub SetDocBackColor()

    On Error Resume Next
    DoEvents
    wbPreview.Document.body.bgColor = GetRGB(cmdColor.Tag)

End Sub

Private Sub wbPreview_NavigateComplete2(ByVal pDisp As Object, URL As Variant)

    On Error Resume Next

    SetDocBackColor

    If URL = "http:///" Or URL = "about:blank" Then Exit Sub

End Sub

