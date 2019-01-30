VERSION 5.00
Object = "{683EFEE4-F060-4ADB-B94E-96A91D02F8D0}#1.0#0"; "FormShaper.ocx"
Begin VB.Form frmNag 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   ClientHeight    =   5088
   ClientLeft      =   900
   ClientTop       =   2160
   ClientWidth     =   7188
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MousePointer    =   11  'Hourglass
   ScaleHeight     =   5088
   ScaleWidth      =   7188
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmrReveal 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1965
      Top             =   3465
   End
   Begin VB.Timer tmrPBAnim 
      Interval        =   10
      Left            =   5895
      Top             =   3675
   End
   Begin VB.PictureBox picPBSrc 
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   3555
      Picture         =   "frmNag.frx":0000
      ScaleHeight     =   156
      ScaleWidth      =   1680
      TabIndex        =   7
      Top             =   3060
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.PictureBox picPB 
      AutoRedraw      =   -1  'True
      Height          =   210
      Left            =   3270
      ScaleHeight     =   168
      ScaleWidth      =   2388
      TabIndex        =   6
      Top             =   2145
      Width           =   2430
   End
   Begin xfxFormShaper.FormShaper fsCtrl 
      Left            =   3600
      Top             =   3330
      _ExtentX        =   1863
      _ExtentY        =   1291
   End
   Begin VB.Timer tmrClose 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   2985
      Top             =   3780
   End
   Begin VB.Label lblDEVer 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "<developers_edition>"
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   180
      Left            =   1905
      TabIndex        =   5
      Top             =   1575
      Width           =   2100
   End
   Begin VB.Label lblUCode 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Unlock Code"
      Height          =   195
      Left            =   390
      TabIndex        =   4
      Top             =   2970
      UseMnemonic     =   0   'False
      Width           =   885
   End
   Begin VB.Label lblmsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program is licensed to"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   195
      TabIndex        =   3
      Top             =   2160
      UseMnemonic     =   0   'False
      Width           =   1890
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Loading..."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Left            =   3270
      TabIndex        =   2
      Top             =   1965
      UseMnemonic     =   0   'False
      Width           =   585
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      Height          =   195
      Left            =   390
      TabIndex        =   1
      Top             =   2685
      UseMnemonic     =   0   'False
      Width           =   675
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   195
      Left            =   390
      TabIndex        =   0
      Top             =   2505
      UseMnemonic     =   0   'False
      Width           =   780
   End
End
Attribute VB_Name = "frmNag"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cx As Long
Dim cy As Long
Dim IsDragging As Boolean
Dim WinAlpha As Integer

'Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
'Private Const conHwndTopmost = -1
'Private Const conHwndNoTopmost = -2
'Private Const conSwpNoActivate = &H10
'Private Const conSwpShowWindow = &H40

Private Sub Form_Load()

    'If IsWinXP Then
    '    WinAlpha = 0
    '    MakeTransparent Me.hWnd, WinAlpha
    '    tmrReveal.Enabled = True
    'End If
    Set Me.Picture = LoadPicture(AppPath + "rsc\nag.gif")
    
    Width = 402 * 15 'Screen.TwipsPerPixelX
    Height = 248 * 15 'Screen.TwipsPerPixelY
    
    CenterForm Me

    NagScreenIsVisible = True
    fsCtrl.ShapeItImage
    'SetWindowPos Me.hwnd, conHwndTopmost, Left / 15, Top / 15, Width / 15, Height / 15, conSwpNoActivate Or conSwpShowWindow
    
    LocalizeUI
    
    lblDEVer.Caption = GetAppTypeName
    
    lblUserName.Caption = USER
    lblCompany.Caption = COMPANY
    lblUCode.Caption = USERSN
    
    lblInfo.Caption = GetLocalizedStr(285)
    
    lblInfo.Top = picPB.Top - lblInfo.Height - 15
    
    On Error Resume Next
    
    picPB.Width = picPB.Width - picPB.Width * (15 / Screen.TwipsPerPixelX - 1)
    picPB.Height = picPB.Height - picPB.Height * (15 / Screen.TwipsPerPixelY - 1)
    
    Dim c As Control
    For Each c In Me.Controls
        c.Left = c.Left - c.Left * (15 / Screen.TwipsPerPixelX - 1)
        c.Top = c.Top - c.Top * (15 / Screen.TwipsPerPixelY - 1)
    Next c
    
    Me.AutoRedraw = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsDragging = True
    cx = x
    cy = y
    MousePointer = vbSizeAll

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If IsDragging Then
        Move Left + (x - cx), Top + (y - cy)
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsDragging = False
    MousePointer = vbDefault

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    'VanishMe Me
    NagScreenIsVisible = False

End Sub

Private Sub lblInfo_Change()

    lblInfo.Refresh

End Sub

Friend Sub tmrClose_Timer()

    If Not InitDone Then Exit Sub
    frmMain.Enabled = True
    Unload Me

End Sub

Private Sub LocalizeUI()

    lblmsg.Caption = GetLocalizedStr(633)

End Sub

Friend Sub tmrPBAnim_Timer()

    Static Xos As Long
    
    '    picPB.PaintPicture picPBSrc, Xos * Screen.TwipsPerPixelX, 0
    '    picPB.Refresh
    '
    '    Xos = Xos - 1
    '    If Xos = -17 Then Xos = 0
    
    picPB.Cls
    picPB.PaintPicture picPBSrc, (Xos - 16) * Screen.TwipsPerPixelX, 0
    picPB.Refresh
    
    Xos = Xos + 1
    If Xos = 17 Then Xos = 0
    
End Sub

'Private Sub tmrReveal_Timer()
'
'    WinAlpha = WinAlpha + 5
'    If WinAlpha >= 256 Then
'        tmrReveal.Enabled = False
'    Else
'        MakeTransparent Me.hWnd, WinAlpha
'    End If
'
'End Sub
