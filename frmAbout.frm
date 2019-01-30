VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "About MyApp"
   ClientHeight    =   6030
   ClientLeft      =   7050
   ClientTop       =   4080
   ClientWidth     =   6945
   ClipControls    =   0   'False
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
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox imgR 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   5835
      MouseIcon       =   "frmAbout.frx":2CFA
      MousePointer    =   99  'Custom
      ScaleHeight     =   750
      ScaleWidth      =   690
      TabIndex        =   19
      Top             =   2580
      Width           =   690
   End
   Begin VB.Timer tmrPrepareEEGG 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   225
      Top             =   3480
   End
   Begin VB.TextBox txtSN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000011&
      Height          =   315
      Left            =   1095
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3945
      Width           =   4050
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   51
      Left            =   240
      Top             =   2055
   End
   Begin VB.PictureBox picScrollContainer 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   1050
      ScaleHeight     =   1695
      ScaleWidth      =   4125
      TabIndex        =   4
      Top             =   1215
      Width           =   4125
      Begin VB.PictureBox picScroll 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ClipControls    =   0   'False
         ForeColor       =   &H80000008&
         Height          =   18525
         Left            =   30
         ScaleHeight     =   18525
         ScaleWidth      =   4065
         TabIndex        =   5
         Top             =   -1715
         Width           =   4065
         Begin VB.TextBox txtMsg 
            Appearance      =   0  'Flat
            BackColor       =   &H8000000F&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   13365
            Left            =   300
            Locked          =   -1  'True
            MousePointer    =   1  'Arrow
            MultiLine       =   -1  'True
            TabIndex        =   6
            Text            =   "frmAbout.frx":2E4C
            Top             =   3795
            Width           =   3240
         End
         Begin VB.Label lblCopyright 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "©2000 All Rights Reserved"
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
            Left            =   900
            TabIndex        =   13
            Top             =   18315
            Width           =   1950
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Xavier Flix"
            Height          =   210
            Left            =   300
            TabIndex        =   12
            Top             =   2145
            UseMnemonic     =   0   'False
            Width           =   915
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Project coordinator & original concept"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   11
            Top             =   2535
            UseMnemonic     =   0   'False
            Width           =   3225
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Xavier Flix"
            Height          =   210
            Left            =   300
            TabIndex        =   10
            Top             =   2790
            UseMnemonic     =   0   'False
            Width           =   915
         End
         Begin VB.Image Image1 
            Height          =   240
            Left            =   0
            Picture         =   "frmAbout.frx":315A
            Stretch         =   -1  'True
            Top             =   1860
            Width           =   240
         End
         Begin VB.Image Image2 
            Height          =   240
            Left            =   0
            Picture         =   "frmAbout.frx":359C
            Stretch         =   -1  'True
            Top             =   2505
            Width           =   240
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Programming"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   9
            Top             =   1890
            UseMnemonic     =   0   'False
            Width           =   1155
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Special Thanks..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   8
            Top             =   3540
            Width           =   1395
         End
         Begin VB.Image Image6 
            Height          =   240
            Left            =   0
            Picture         =   "frmAbout.frx":39DE
            Stretch         =   -1  'True
            Top             =   3517
            Width           =   240
         End
         Begin VB.Image Image7 
            Height          =   1185
            Left            =   945
            Stretch         =   -1  'True
            Top             =   17235
            Width           =   1800
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "®"
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
            Left            =   2475
            TabIndex        =   7
            Top             =   17955
            Width           =   150
         End
      End
   End
   Begin VB.TextBox txtReg 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000011&
      Height          =   495
      Left            =   1095
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   3420
      Width           =   4050
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   5550
      TabIndex        =   2
      Top             =   5535
      Width           =   1260
   End
   Begin VB.PictureBox picLogo 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1155
      Left            =   45
      ScaleHeight     =   1155
      ScaleWidth      =   6855
      TabIndex        =   0
      Top             =   45
      Width           =   6855
      Begin VB.Label lblDEVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "<developers_edition>"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   180
         Left            =   1695
         TabIndex        =   1
         Top             =   840
         UseMnemonic     =   0   'False
         Width           =   2100
      End
   End
   Begin VB.Timer tmrStartScroll 
      Interval        =   3500
      Left            =   330
      Top             =   2730
   End
   Begin VB.Image Image5 
      Height          =   480
      Left            =   3330
      Picture         =   "frmAbout.frx":3E20
      ToolTipText     =   "Presets Support"
      Top             =   5415
      Width           =   480
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   2625
      Picture         =   "frmAbout.frx":4AEA
      ToolTipText     =   "AddIns Support"
      Top             =   5415
      Width           =   480
   End
   Begin VB.Image imgRORLogo 
      Height          =   570
      Left            =   1785
      Picture         =   "frmAbout.frx":57B4
      ToolTipText     =   "ROR Import Support"
      Top             =   5370
      Width           =   615
   End
   Begin VB.Image imgSEOLogo 
      Appearance      =   0  'Flat
      Height          =   570
      Left            =   180
      Picture         =   "frmAbout.frx":5E42
      ToolTipText     =   "Search Engine Optimizations"
      Top             =   5370
      Width           =   1410
   End
   Begin VB.Label lblResInfo 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5670
      TabIndex        =   20
      Top             =   3210
      UseMnemonic     =   0   'False
      Width           =   645
   End
   Begin VB.Image imgSIMONLogo 
      Height          =   1185
      Left            =   5160
      Stretch         =   -1  'True
      Top             =   1275
      Width           =   1800
   End
   Begin VB.Label lblEngineVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Engine Version"
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
      Left            =   5760
      TabIndex        =   15
      Top             =   4065
      UseMnemonic     =   0   'False
      Width           =   1050
   End
   Begin VB.Image Image3 
      Height          =   240
      Left            =   795
      Picture         =   "frmAbout.frx":695E
      Stretch         =   -1  'True
      Top             =   3105
      Width           =   240
   End
   Begin VB.Label lblRegMsg 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This program is registered to"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1095
      TabIndex        =   18
      Top             =   3135
      UseMnemonic     =   0   'False
      Width           =   2460
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Version"
      Height          =   210
      Left            =   6210
      TabIndex        =   17
      Top             =   3840
      UseMnemonic     =   0   'False
      Width           =   600
   End
   Begin VB.Label lblDisclaimer 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":6DA0
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
      Height          =   570
      Left            =   180
      TabIndex        =   16
      Top             =   4500
      UseMnemonic     =   0   'False
      Width           =   6435
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim IsALTPressed As Boolean
Dim cx As Long
Dim cy As Long
Dim IsDragging As Boolean

Private WithEvents msgSubClass As xfxSC
Attribute msgSubClass.VB_VarHelpID = -1

Private Sub cmdOK_Click()

    Unload Me

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    On Error Resume Next

    Select Case KeyCode
        Case 107 '+
            tmrScroll.Interval = tmrScroll.Interval - 25
        Case 109 '-
            tmrScroll.Interval = tmrScroll.Interval + 25
    End Select

End Sub

Private Sub Form_Load()

    On Error GoTo Form_Load_Error

    Set picLogo.Picture = LoadPicture(AppPath + "rsc\about.jpg")
    Set imgSIMONLogo.Picture = LoadPicture(AppPath + "rsc\xfxlogo.gif")
    Set Image7.Picture = LoadPicture(AppPath + "rsc\xfxlogo.gif")

    'SetupCharset Me
    CenterForm Me
    
    LocalizeUI
    
    Set msgSubClass = New xfxSC
    msgSubClass.SubClassHwnd Me.hwnd, True

    Me.caption = "About DHTML Menu Builder" + IIf(IsDEMO, " DEMO", vbNullString)
    
    lblDEVer.caption = GetAppTypeName
    
    lblVersion.caption = DMBVersion
    lblEngineVersion.caption = "Engine " + EngineVersion
    
    txtReg.Text = USER + vbCrLf + COMPANY
    txtSN.Text = Str2HEX(USERSN)
    txtMsg.Text = "To my wife and kids... for their patience!" + vbCrLf + vbCrLf + _
             "To Vikki Dawson for her great support, for the excellent review about version 1.5 and for accepting the challenge to try the first 2.0 beta!" + vbCrLf + vbCrLf + _
             "To Lyon L. for accepting the challenge to try the first 2.0 beta and creating a very nice review about version 2.8" + vbCrLf + vbCrLf + _
             "To Robert Lee for his unconditional support and help!" + vbCrLf + vbCrLf + _
             "To all of you who bought version 1.5!" + vbCrLf + "Thank you!" + vbCrLf + vbCrLf + _
             "To Jonathan Hilgeman for designing the new logo used on versions 2.6 and above." + vbCrLf + "Excellent work!" + vbCrLf + vbCrLf + _
             "To Pai Yili for being a great friend and helpful partner." + vbCrLf + "Thanks pal!" + vbCrLf + vbCrLf + _
             "To my wife again... for being so patient and dedicated to help me in ways beyond anyone could ever wish for..." + vbCrLf + "Love ya!" + vbCrLf + vbCrLf + _
             "To 'BS' for being the first company that ever bought a program from me." + vbCrLf + vbCrLf + _
             "To Henri for his unconditional help and willingness." + vbCrLf + "A huge part of DMB is also yours..." + vbCrLf + vbCrLf + _
             "To Pai, Derrick, Mark S., Matt, Robert, Neil and to all of you that kept me updated on bugs and problems on the forum. Because without them, DMB 3 would be a program full of bugs :)" + vbCrLf + vbCrLf + _
             "To kwayczar, for the new border preview on the groups, which is now also being used on the Live Preview component." + vbCrLf + vbCrLf + _
             "To Bill Lanides from Intrasphere, for giving us a simple solution to a very old problem. His solution improves the menus running under DOM-compliant browsers considerably." + vbCrLf + vbCrLf + _
             "To RegNet for their amazing support!" + vbCrLf + vbCrLf + _
             "To Henri, Christian and Carmen for helping in the translation packages." + vbCrLf + vbCrLf + _
             "...and last but not least, to all of you that have helped me in some way or another improve DHTML Menu Builder"
    
    lblCopyright.caption = "©" & Year(Date) & " All Rights Reserved"
    If Not IsDEMO Then
        txtMsg.Text = txtMsg + vbCrLf + vbCrLf + vbCrLf + "And since you've been waiting this long, huge kudos to our favorite customer: " + USER + "!!!"
    End If
    
    picScroll.Top = -1715
    
    If ResellerID = "" Then
        imgR.Visible = False
        lblResInfo.Visible = False
    Else
        With imgR
            Set .Picture = LoadPicture(AppPath + "rinfo\" + ResellerID + "\logo.gif")
            DoEvents
            .Left = imgSIMONLogo.Left + imgSIMONLogo.Width / 2 - .Width / 2
            .Top = imgSIMONLogo.Top + imgSIMONLogo.Height + 2 * Screen.TwipsPerPixelX
            .BackColor = IIf(ResellerInfo(6) = -1, Me.BackColor, ResellerInfo(6))
            .ToolTipText = "Click here to visit " + ResellerInfo(0)
        End With
        With lblResInfo
            .caption = ResellerInfo(8)
            .Left = imgR.Left - 4 * Screen.TwipsPerPixelX
            .Top = imgR.Top + imgR.Height '+ 1 * Screen.TwipsPerPixelY
            .Width = imgR.Width + 8 * Screen.TwipsPerPixelX
            .BackColor = IIf(ResellerInfo(6) = -1, Me.BackColor, ResellerInfo(6))
            .ForeColor = IIf(ResellerInfo(7) = -1, Me.ForeColor, ResellerInfo(7))
            .Height = 405 + 18 * 15
        End With
    End If
    
    'RevealMe Me

    On Error GoTo 0
    Exit Sub

Form_Load_Error:

    MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Form.frmAbout.Form_Load"
    
End Sub

Private Sub DoEEGG()

    SaveFile SimonFile, modZlib.UnCompress(AppPath + "rsc\rsc.dat")
    
    Shell SimonFile, vbNormalFocus

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    msgSubClass.SubClassHwnd Me.hwnd, False
    'VanishMe Me

End Sub

Private Sub Image7_DblClick()

    If Not IsDEMO Then
        If IsALTPressed Then
            imgR.Visible = False
            lblResInfo.Visible = False
            With imgSIMONLogo
                .Picture = LoadResPicture(101, vbResIcon)
                .Width = 1500
                .Height = 1500
            End With
            With Image7
                .Width = 1300
                .Height = 1600
                .Picture = LoadResPicture(101, vbResBitmap)
            End With
            Label10.Visible = False
            lblCopyright.Visible = False
            tmrPrepareEEGG.Enabled = True
        End If
    End If

End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsALTPressed = (Shift = vbAltMask)

End Sub

Private Sub imgR_Click()

    RunShellExecute "Open", ResellerInfo(1), 0, 0, 0

End Sub

Private Sub msgSubClass_NewMessage(ByVal hwnd As Long, uMsg As Long, wParam As Long, lParam As Long, Cancel As Boolean, UseRetVal As Boolean, RetVal As Long)

    Select Case hwnd
        Case Me.hwnd
            Select Case uMsg
                Case WM_PAINT
                    DrawBorder
            End Select
    End Select

End Sub

Private Sub DrawBorder()

    Dim r As RECT
    
    Cls
    
    With r
        .Left = 0
        .Top = 0
        .Right = Width / Screen.TwipsPerPixelX
        .bottom = Height / Screen.TwipsPerPixelY
    End With
    DrawEdge Me.hDc, r, BDR_RAISED, BF_RECT
    
    With r
        .Left = lblRegMsg.Left / Screen.TwipsPerPixelX
        .Top = lblRegMsg.Top / Screen.TwipsPerPixelY - 15
        .Right = .Left + picScrollContainer.Width / Screen.TwipsPerPixelX
        .bottom = .Top + 2
    End With
    DrawEdge Me.hDc, r, BDR_SUNKENOUTER, BF_RECT
    
    With r
        .Left = 8
        .Top = lblDisclaimer.Top / Screen.TwipsPerPixelY - 8
        .Right = Width / Screen.TwipsPerPixelX - 8
        .bottom = .Top + 2
    End With
    DrawEdge Me.hDc, r, BDR_SUNKENOUTER, BF_RECT
    
    With r
        .Top = (lblDisclaimer.Top + lblDisclaimer.Height) / Screen.TwipsPerPixelY + 8
        .bottom = .Top + 2
    End With
    DrawEdge Me.hDc, r, BDR_SUNKENOUTER, BF_RECT
    
    With r
        .Top = (imgSEOLogo.Top / Screen.TwipsPerPixelY) + 2
        .bottom = ((imgSEOLogo.Top + imgSEOLogo.Height) / Screen.TwipsPerPixelY) - 2
        .Left = 112
        .Right = .Left + 2
    End With
    DrawEdge Me.hDc, r, BDR_SUNKENOUTER, BF_RECT
    
    With r
        .Left = 167
        .Right = .Left + 2
    End With
    DrawEdge Me.hDc, r, BDR_SUNKENOUTER, BF_RECT
    
    With r
        .Left = 214
        .Right = .Left + 2
    End With
    DrawEdge Me.hDc, r, BDR_SUNKENOUTER, BF_RECT

End Sub

Private Sub picLogo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsDragging = True
    cx = x
    cy = y
    MousePointer = vbSizeAll

End Sub

Private Sub picLogo_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    If IsDragging Then
        Move Left + (x - cx), Top + (y - cy)
    End If

End Sub

Private Sub picLogo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    IsDragging = False
    MousePointer = vbDefault

End Sub

Private Sub tmrPrepareEEGG_Timer()

    tmrPrepareEEGG.Enabled = False
    DoEEGG

End Sub

Private Sub tmrScroll_Timer()

    picScroll.Top = picScroll.Top - Screen.TwipsPerPixelY
    If picScroll.Top <= -17130 Then tmrScroll.Enabled = False
    
    If Abs(picScroll.Top) Mod 100 = 0 Then
        picScroll.Refresh
    End If

End Sub

Private Sub tmrStartScroll_Timer()

    tmrStartScroll.Enabled = False
    tmrScroll.Enabled = True

End Sub

Private Sub txtMsg_Click()

    cmdOK.SetFocus

End Sub

Private Sub LocalizeUI()

    lblRegMsg.caption = GetLocalizedStr(633)

End Sub
