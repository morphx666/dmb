VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{2A2AD7CA-AC77-46F3-84DC-115021432312}#1.0#0"; "HREF.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmNewReg 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "DHTML Menu Builder Registration"
   ClientHeight    =   5355
   ClientLeft      =   7440
   ClientTop       =   5550
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNewReg.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5565
   ShowInTaskbar   =   0   'False
   Begin href.uchref1 uchLostRegInfo 
      Height          =   315
      Left            =   195
      TabIndex        =   22
      Top             =   3225
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   556
      Caption         =   "Have you lost your registration details?"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "Tahoma"
      FontSize        =   8.25
      ForeColor       =   16711680
      URL             =   "http://software.xfx.net/utilities/dmbuilder/lostpwd.php"
   End
   Begin VB.Timer tmrAutoValidate 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2100
      Top             =   4875
   End
   Begin VB.Timer tmrSeconds 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   4875
   End
   Begin VB.PictureBox picStatus 
      BorderStyle     =   0  'None
      Height          =   1005
      Left            =   180
      ScaleHeight     =   1005
      ScaleWidth      =   3585
      TabIndex        =   8
      Top             =   3630
      Width           =   3585
      Begin MSComctlLib.ProgressBar pbTimeOut 
         Height          =   180
         Left            =   300
         TabIndex        =   15
         Top             =   825
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   318
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Max             =   90
      End
      Begin VB.Label lblStep1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Verifying Internet Connection"
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
         Left            =   300
         TabIndex        =   10
         Top             =   60
         Width           =   2130
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   45
         TabIndex        =   9
         Top             =   45
         Width           =   195
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   45
         TabIndex        =   13
         Top             =   585
         Width           =   195
      End
      Begin VB.Label lblStep3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Validating Registration Information"
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
         Left            =   300
         TabIndex        =   14
         Top             =   600
         Width           =   2490
      End
      Begin VB.Label lblStep2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Preparing Data"
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
         Left            =   300
         TabIndex        =   12
         Top             =   330
         Width           =   1080
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Wingdings"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   45
         TabIndex        =   11
         Top             =   315
         Width           =   195
      End
   End
   Begin xfxLine3D.ucLine3D uc3DLine2 
      Height          =   30
      Left            =   30
      TabIndex        =   7
      Top             =   3585
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   53
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   360
      Left            =   45
      TabIndex        =   20
      Top             =   4935
      Width           =   1170
   End
   Begin VB.CommandButton cmdRegister 
      Caption         =   "Register"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4350
      TabIndex        =   18
      Top             =   4335
      Width           =   1170
   End
   Begin VB.CommandButton cmdValidate 
      Caption         =   "Validate"
      Height          =   360
      Left            =   4350
      TabIndex        =   17
      Top             =   3870
      Width           =   1170
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Height          =   360
      Left            =   4350
      TabIndex        =   21
      Top             =   4935
      Width           =   1170
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   30
      TabIndex        =   19
      Top             =   4785
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   53
   End
   Begin VB.TextBox txtCompany 
      Height          =   315
      Left            =   225
      TabIndex        =   6
      Top             =   2865
      Width           =   3855
   End
   Begin VB.TextBox txtName 
      Height          =   315
      Left            =   225
      TabIndex        =   4
      Top             =   2280
      Width           =   3855
   End
   Begin VB.TextBox txtOrder 
      Height          =   315
      Left            =   225
      TabIndex        =   2
      Top             =   1695
      Width           =   3855
   End
   Begin SHDocVwCtl.WebBrowser wbCtrl 
      Height          =   135
      Left            =   1650
      TabIndex        =   16
      Top             =   5025
      Width           =   135
      ExtentX         =   238
      ExtentY         =   238
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   1
      RegisterAsDropTarget=   0
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
      Location        =   ""
   End
   Begin VB.Label lblInfo 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNewReg.frx":0CCA
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   75
      TabIndex        =   0
      Top             =   975
      Width           =   5205
   End
   Begin VB.Image imgLogo 
      Height          =   900
      Left            =   0
      Top             =   0
      Width           =   6015
   End
   Begin VB.Label lblCompany 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      Height          =   210
      Left            =   225
      TabIndex        =   5
      Top             =   2640
      Width           =   750
   End
   Begin VB.Label lblUserName 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      Height          =   210
      Left            =   225
      TabIndex        =   3
      Top             =   2055
      Width           =   885
   End
   Begin VB.Label lblOrderNumber 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Order Number"
      Height          =   210
      Left            =   225
      TabIndex        =   1
      Top             =   1470
      Width           =   1170
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "mnuHelp"
      Begin VB.Menu mnuHelpRegHelp 
         Caption         =   "Registration Help..."
      End
      Begin VB.Menu mnuHelpPrivacyInfo 
         Caption         =   "Privacy Information..."
      End
   End
End
Attribute VB_Name = "frmNewReg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsMD5 As MD5
Dim ForceExit As Boolean
Dim vServerAddress As String

Dim TimeoutValue As Integer

'http://216.55.183.238/dmbkeygen/dmbkeygen.asp/dmbkg.cgi?User=Xavier%20Flix%20S.&Company=xFX%20JumpStart%20:%20Software%20Division&Order=xavier%40xfx.net&vKey=A6450BF47D06FDA402919F897ED109AB&DMBVer=4.20.001&Mode=Application&wd=6&rseed=42901.09&appType=STD

Private Sub CtrlsState(State As Boolean)

    cmdValidate.Enabled = State
    cmdRegister.Enabled = State
    cmdClose.Enabled = State
    
    txtOrder.Enabled = State
    txtName.Enabled = State
    txtCompany.Enabled = State

End Sub

Private Sub cmdClose_Click()

    If Not DoSilentValidation Then
        If cmdRegister.Enabled Then
            If MsgBox(GetLocalizedStr(654), vbYesNo + vbQuestion, GetLocalizedStr(628)) = vbNo Then
                Exit Sub
            End If
        End If
    End If
    
    USER = "DEMO"
    COMPANY = "DEMO"

    Unload Me

End Sub

Private Sub cmdHelp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    PopupMenu mnuHelp, , cmdHelp.Left - 15, cmdHelp.Top + cmdHelp.Height

End Sub

Private Sub cmdRegister_Click()

    SaveSetting App.EXEName, "RegInfo", "OrderNum", txtOrder.Text

    USER = txtName.Text
    COMPANY = txtCompany.Text
    
    Dim exMsg As String
    If IsWinVistaOrAbove Then
        exMsg = vbCrLf + vbCrLf + "You will now see a dialog box asking you to allow an application (register.exe) access to your computer; make sure you select ""Allow"" so that DHTML Menu Builder can properly complete the registration process."
    End If
    
    MsgBox "DHTML Menu Builder will now restart." + vbCrLf + "Please be patient as this process can take some time." + exMsg, vbInformation + vbOKOnly, "Registration"
    
    Unload Me

End Sub

Private Function FixRegParam(ByVal txt As String) As String

    txt = Trim(txt)
    txt = Replace(txt, vbLf, "")
    txt = Replace(txt, vbCr, "")
    txt = Replace(txt, "  ", " ")
    
    FixRegParam = txt

End Function

Private Sub cmdValidate_Click()

    DoValidate

End Sub

Friend Sub DoValidate()

          Dim vKey As String
          Dim md As String
          Dim IsOnline As Boolean
          Dim appType As String
          
10        On Error GoTo DoValidate_Error

20        ResetStatus
          
30        CtrlsState False
40        MousePointer = vbHourglass
          
50        pbTimeOut.Value = 0
60        vServerAddress = ""
          
70        DoEvents
          
80        txtOrder.Text = FixRegParam(Replace(txtOrder.Text, "#", ""))
90        txtName.Text = FixRegParam(txtName.Text)
100       txtCompany.Text = FixRegParam(txtCompany.Text)
          
110       IsOnline = IsConnectedToInternet
120       lblStatus(0).caption = IIf(IsOnline, Chr(Hex2Dec("FE")), Chr(Hex2Dec("FD")))
130       If Not IsOnline Then GoTo subForceExit
          
140       TimeoutValue = CInt(PingUDP("xfx.net") * 1.5)
150       If TimeoutValue < 120 Then TimeoutValue = 120
160       pbTimeOut.Max = TimeoutValue
          
170       Set clsMD5 = New MD5
180       vKey = txtOrder.Text + txtName.Text + txtCompany.Text
190       lblStatus(1).caption = Chr(Hex2Dec("A8"))
200       md = clsMD5.DigestFileToHexStr(AppPath + "register.exe")
210       lblStatus(1).caption = IIf(LenB(md) <> 0, Chr(Hex2Dec("FE")), Chr(Hex2Dec("FD")))
220       If LenB(md) = 0 Then GoTo subForceExit

230       If LenB(txtOrder.Text) = 0 Or LenB(txtName.Text) = 0 Then
240           If Not DoSilentValidation Then MsgBox GetLocalizedStr(640), vbInformation + vbOKOnly, GetLocalizedStr(641)
250           CtrlsState True
260           cmdRegister.Enabled = False
270           MousePointer = vbDefault
280       Else
290           If (Len(txtOrder.Text) >= 28 And InStr(txtOrder.Text, "@") = 0) Then
300               If Not DoSilentValidation Then MsgBox GetLocalizedStr(651), vbInformation + vbOKOnly, GetLocalizedStr(641)
310               CtrlsState True
320               cmdRegister.Enabled = False
330               MousePointer = vbDefault
340           Else
350               lblStatus(2).caption = Chr(Hex2Dec("A8"))
360               ForceExit = False
370               tmrSeconds.Enabled = True
380               If CheckHOST Then
390                   If IsDebug Then SaveSetting App.EXEName, "RegInfo", "vServerAddress", vServerAddress
                      
                #If DEVVER = 1 Then
400                       appType = "DEV"
                #Else
                    #If LITE = 1 Then
410                           appType = "LIT"
                    #Else
420                           appType = "STD"
                    #End If
                #End If
430                   If appType = "DEV" Then appType = "STD"
                      
                      'vServerAddress = "64.77.42.34/dmbkeygen/dmbkeygen.asp/dmbkgbeta.cgi"
440                   wbCtrl.Navigate "http://" + vServerAddress + "?" + _
                                  "User=" + xFixURL(txtName.Text) + _
                                  "&Company=" + xFixURL(txtCompany.Text) + _
                                  "&Order=" + xFixURL(txtOrder.Text) + _
                                  "&vKey=" + md + _
                                  "&DMBVer=" + DMBVersion + _
                                  "&Mode=Application" + _
                                  "&wd=" & CStr(Weekday(Date, vbMonday)) + _
                                  "&rseed=" & Timer & _
                                  "&appType=" + appType
450               Else
460                   If Not DoSilentValidation Then
470                       MsgBox "An error has occurred while contacting the validation server." + vbCrLf + _
                              "If you're behind a firewall try enabling access through the Port 80 or temporarily disable the firewall while you validate DHTML Menu Builder.", vbCritical, "Error Validating"
480                       ResetControls
490                       GoTo subForceExit
500                   End If
510               End If
520           End If
530       End If
          
540       Exit Sub
          
subForceExit:
550       MousePointer = vbDefault
560       cmdClose.Enabled = True

570       On Error GoTo 0
580       Exit Sub

DoValidate_Error:

590       MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Form.frmNewReg.DoValidate"
600       GoTo subForceExit

End Sub

Private Function CheckHOST() As Boolean

    On Error GoTo ExitSub
    
    If InStr(1, Command$, "/vserver:") Then
        vServerAddress = Split(Command$, ":")(1)
        If InStr(vServerAddress, " ") Then vServerAddress = Left(vServerAddress, InStr(vServerAddress, " ") - 1)
        CheckHOST = True
        Exit Function
    End If
    
    wbCtrl.Navigate "https://xfx.net/dmbvs.php?mode=2&seed=" & (Int((Timer - 1) * Rnd + Timer))
    
    While LenB(vServerAddress) = 0 And (Not ForceExit)
        DoEvents
    Wend
    
    If ForceExit Then
        tmrSeconds.Enabled = False
        cmdClose.Enabled = True
    End If
    
    CheckHOST = (LenB(vServerAddress) <> 0) And (Not ForceExit)
        
    Exit Function
    
ExitSub:
    CheckHOST = True

End Function

Private Function xFixURL(ByVal s As String) As String

    s = EscapePath(s)
    s = Replace(s, "&", "%26")
    s = Replace(s, "@", "%40")
    
    xFixURL = Trim(s)

End Function

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    
    If KeyCode = 27 Then ForceExit = True
    
End Sub

Private Sub Form_Load()

    Dim rUser As String
    Dim rCompany As String
    Dim dcs As Integer
    
    If DoSilentValidation And IsWinXP Then MakeTransparent Me.hwnd, 0
    
    mnuHelp.Visible = False
    Height = 5075 + GetClientTop(Me.hwnd)
    
    dcs = lblStatus(0).Font.Charset
    
    imgLogo.Picture = LoadPicture(AppPath + "rsc\unwxfxlogo.jpg")
    
    LocalizeUI
    SetupCharset Me
    CenterForm Me
    DisableCloseButton Me
    
    lblStatus(0).Font.Charset = dcs
    lblStatus(1).Font.Charset = dcs
    lblStatus(2).Font.Charset = dcs
    
    ResetStatus
    
    rUser = GetSetting(App.EXEName, "RegInfo", "User", "DEMO")
    rCompany = GetSetting(App.EXEName, "RegInfo", "Company", "DEMO")
    
    If rUser <> "DEMO" And rCompany <> "DEMO" Then
        txtName.Text = rUser
        txtCompany.Text = rCompany
        txtOrder.Text = GetSetting(App.EXEName, "RegInfo", "OrderNum", "")
    End If
    
    If DoSilentValidation Then
        Me.Visible = False
        txtOrder.Text = ORDERNUMBER
        txtCompany.Text = COMPANY
        txtName.Text = USER
        tmrAutoValidate.Enabled = True
    End If

End Sub

Private Sub ResetStatus()

    lblStatus(0).caption = Chr(Hex2Dec("9F"))
    lblStatus(1).caption = Chr(Hex2Dec("9F"))
    lblStatus(2).caption = Chr(Hex2Dec("9F"))

End Sub

Private Sub mnuHelpPrivacyInfo_Click()

    showHelp "privacyinfo.htm"

End Sub

Private Sub mnuHelpRegHelp_Click()

    showHelp "reginfo.htm"

End Sub

Private Sub tmrAutoValidate_Timer()

    tmrAutoValidate.Enabled = False
    DoValidate
    
End Sub

Private Sub tmrSeconds_Timer()

    If (pbTimeOut.Value >= TimeoutValue) Or (pbTimeOut.Value >= pbTimeOut.Max) Then
        MsgBox "The connection has timedout: " & TimeoutValue, vbCritical, "Error Connecting"
        ForceExit = True
        tmrSeconds.Enabled = False
        cmdClose.Enabled = True
    Else
        pbTimeOut.Value = pbTimeOut.Value + 1
    End If

End Sub

Private Function encDate(Seed As Integer) As String

    encDate = CStr(Int(10000 * Sqr(Weekday(Date, vbMonday) * Seed)))

End Function

Private Sub txtCompany_Change()

    ValidateRegInfo

End Sub

Private Sub txtName_Change()

    ValidateRegInfo

End Sub

Private Sub txtOrder_Change()

    ValidateRegInfo

End Sub

Private Sub ValidateRegInfo()

    On Error GoTo ExitSub
    
    cmdValidate.Enabled = True
    
    If chkTxt("522720") Or _
        chkTxt("492276") Or _
        chkTxt("F559-4JRP-05W7-7U") Or _
        chkTxt("9142-7821-6479-XT57") Or _
        chkTxt("4N96-3080-1604-NC31") Or _
        chkTxt("70984325213") Or _
        chkTxt("3323-8788-3150-GJ72") Or _
        chkTxt("4N96-3080-1604-NC31") Or _
        chkTxt("Debra Smith") Or _
        chkTxt("Ramon Alvarez") Or _
        chkTxt("king_KINK") Or _
        chkTxt("DEVON MICHAELS") Or _
        chkTxt("devonsbodyshop") Or _
        chkTxt("4598460E4877439CC6A1E952AAEC40F6") Or _
        chkTxt("719530") Or _
        chkTxt("772687") Or _
        chkTxt("TEAMDVT") Or _
        chkTxt("19A0D083F8AEEA1667D20153D141A5A57B61") Or _
        chkTxt("Re@ldaTa") Then
        cmdValidate.Enabled = False
        Exit Sub
    End If
       
ExitSub:
    If Err.number <> 0 Then cmdValidate.Enabled = True

End Sub

Private Function chkTxt(txt As String) As Boolean

    chkTxt = (InStr(1, txtOrder.Text, txt, vbTextCompare) <> 0) Or _
            (InStr(1, txtName.Text, txt, vbTextCompare) <> 0) Or _
            (InStr(1, txtCompany.Text, txt, vbTextCompare) <> 0)

End Function

Private Sub uchLostRegInfo_Click()

    RunShellExecute "Open", uchLostRegInfo.url, 0, 0, 0

End Sub

Private Sub wbCtrl_DocumentComplete(ByVal pDisp As Object, url As Variant)

          Dim doc As IHTMLDocument2
          Dim divs As IHTMLElementCollection
          Dim IsValid As Boolean
          Dim sDate As String
          Dim sUrl As String
          Dim sStep As Integer
          
10        On Error GoTo CheckFailed
          
20        sStep = 1
          
30        sUrl = CStr(url)
40        If InStr(1, sUrl, "/dmbvs.php", vbTextCompare) > 0 Then
50            Do While (wbCtrl.Document Is Nothing) And (Not ForceExit)
60                DoEvents
70            Loop
80            Set doc = wbCtrl.Document
90            Set divs = doc.All.tags("div")
100           vServerAddress = divs("sip").innerHTML
110           Exit Sub
120       End If
          
130       sStep = 2
          
140       If InStr(1, sUrl, "/dmbkeygen/dmbkeygen.asp", vbTextCompare) = 0 Then
150           If InStr(1, sUrl, "/vsoffline.asp", vbTextCompare) = 0 Then GoTo ExitSub
160           If InStr(1, sUrl, "http://64.77.42.34/", vbTextCompare) = 0 Then GoTo ExitSub
170       End If
          
180       sStep = 3
          
190       Do While (wbCtrl.Document Is Nothing) And (Not ForceExit)
200           DoEvents
210       Loop
          
220       sStep = 4
          
230       tmrSeconds.Enabled = False
240       pbTimeOut.Value = pbTimeOut.Max
250       If ForceExit Then GoTo CheckFailed
          
260       sStep = 5
          'Debug.Print wbCtrl.Document.body.innerHTML
          
270       Set doc = wbCtrl.Document
280       sStep = 6
290       Set divs = doc.All.tags("div")
300       sStep = 7
310       IsValid = (Split(divs("resp").innerHTML, "|")(0) = "TRUE")
320       sStep = 8
330       USERSN = Split(divs("resp").innerHTML, "|")(1)
340       sStep = 9
350       sDate = Split(divs("resp").innerHTML, "|")(2)
          
360       sStep = 10
          
370       If sDate <> encDate(Len(txtName.Text)) Then
380           USERSN = ""
390           IsValid = False
400       End If
          
410       sStep = 11
          
420       If IsValid Then
430           sStep = 12
440           cmdRegister.Enabled = True
450           If Not DoSilentValidation Then MsgBox GetLocalizedStr(642), vbInformation + vbOKOnly, GetLocalizedStr(643)
460       Else
470           sStep = 13
480           ResetControls
490           If Not DoSilentValidation Then
500               With frmErrValidation
510                   .txtMsg.Text = GetLocalizedStr(644)
520                   .txtServerResponse.Text = USERSN
530                   .Show vbModal
540               End With
550           End If
560       End If
          
570       sStep = 14
          
580       lblStatus(2).caption = IIf(IsValid, Chr(Hex2Dec("FE")), Chr(Hex2Dec("FD")))
          
590       sStep = 15
          
ExitSub:
600       MousePointer = vbDefault
610       cmdClose.Enabled = True
          
620       sStep = 16
          
630       If DoSilentValidation Then
640           If IsValid Then
650               Unload Me
660           Else
670               cmdClose_Click
680           End If
690       End If
          
700       Exit Sub
          
CheckFailed:
          Dim ErrorNumber As String
          Dim ErrorMessage As String
          Dim ErrorLine As String
          Dim errStr As String
          
710       ForceExit = True
          
720       ErrorNumber = Err.number
730       ErrorMessage = Err.Description
740       ErrorLine = Erl
          
750       ResetControls
          
760       On Error Resume Next
          
770       If Not DoSilentValidation Then
780           Do
790               DoEvents
800           Loop While frmErrDlgIsVisible
810           With frmErrValidation
820               .txtMsg.Text = GetLocalizedStr(653)
830               .txtServerResponse.Text = GetLocalizedStr(652)
840               If ErrorNumber <> 0 Then
850                   .txtServerResponse.Text = .txtServerResponse.Text + vbCrLf + "Error " & ErrorNumber & " (" & ErrorLine & "): " & ErrorMessage
860               End If
870               .Show vbModal
880           End With
890       End If
          
900       errStr = Str2HEX(":ERL:" & ErrorLine & ":ERN:" & ErrorNumber & ":ERM:" & ErrorMessage & ":URL:" & sUrl & ":FRE:" & CStr(ForceExit) & ":STP:" & sStep)
910       If doc Is Nothing Then
920           SaveSetting "DMB", "RegInfo", "ServerResponse", "1:No Document:" + errStr
930       Else
940           If doc.body Is Nothing Then
950               SaveSetting "DMB", "RegInfo", "ServerResponse", "2:No Body:" + errStr
960           Else
970               SaveSetting "DMB", "RegInfo", "ServerResponse", "3:DCB:" & doc.body.innerHTML & ":" + errStr
980           End If
990       End If
          
1000      GoTo ExitSub

End Sub

Private Sub ResetControls()

    CtrlsState True
    cmdValidate.Enabled = True
    cmdRegister.Enabled = False

End Sub

Private Sub LocalizeUI()

    lblInfo.caption = GetLocalizedStr(634)
    lblOrderNumber.caption = GetLocalizedStr(635)
    lblUserName.caption = GetLocalizedStr(636)
    lblCompany.caption = GetLocalizedStr(637)
    cmdValidate.caption = GetLocalizedStr(638)
    cmdRegister.caption = GetLocalizedStr(176)
    cmdClose.caption = GetLocalizedStr(424)
    cmdHelp.caption = GetLocalizedStr(197)
    mnuHelpPrivacyInfo.caption = GetLocalizedStr(639) + "..."
    mnuHelpRegHelp.caption = GetLocalizedStr(647) + "..."
    
    lblStep1.caption = GetLocalizedStr(698)
    lblStep2.caption = GetLocalizedStr(700)
    lblStep3.caption = GetLocalizedStr(701)

End Sub
