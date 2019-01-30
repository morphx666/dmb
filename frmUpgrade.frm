VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmUpgrade 
   Caption         =   "Upgrade DHTML Menu Builder"
   ClientHeight    =   6045
   ClientLeft      =   4500
   ClientTop       =   5025
   ClientWidth     =   5970
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmUpgrade.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6045
   ScaleWidth      =   5970
   Begin VB.CommandButton cmdCheck 
      Caption         =   "Check"
      Height          =   375
      Left            =   3705
      TabIndex        =   23
      Top             =   5610
      Width           =   1065
   End
   Begin VB.Timer TimerDoCheck 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1590
      Top             =   5310
   End
   Begin VB.Frame frameLatestVer 
      BorderStyle     =   0  'None
      Caption         =   "Latest Version"
      Height          =   1380
      Left            =   930
      TabIndex        =   7
      Top             =   1710
      Width           =   5670
      Begin VB.Label lblSize 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pP-----"
         Height          =   195
         Left            =   1410
         TabIndex        =   22
         Top             =   1140
         Width           =   480
      End
      Begin VB.Label lblUpdateSize 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Update Size:"
         Height          =   195
         Left            =   450
         TabIndex        =   21
         Top             =   1140
         Width           =   915
      End
      Begin VB.Label lblLatestVerC 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application:"
         Height          =   195
         Left            =   525
         TabIndex        =   8
         Top             =   315
         Width           =   840
      End
      Begin VB.Label lblRelDay 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pP---------------"
         Height          =   195
         Left            =   1410
         TabIndex        =   13
         Top             =   915
         Width           =   1080
      End
      Begin VB.Label lblLatestVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00.000"
         Height          =   195
         Left            =   1410
         TabIndex        =   9
         Top             =   315
         Width           =   660
      End
      Begin VB.Label lblReleased 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Released:"
         Height          =   195
         Left            =   645
         TabIndex        =   12
         Top             =   915
         Width           =   720
      End
      Begin VB.Label lblEngVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00.000"
         Height          =   195
         Left            =   1410
         TabIndex        =   11
         Top             =   585
         Width           =   660
      End
      Begin VB.Label lblLatestVerE 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine:"
         Height          =   195
         Left            =   825
         TabIndex        =   10
         Top             =   585
         Width           =   540
      End
   End
   Begin VB.Frame frameYourVer 
      BorderStyle     =   0  'None
      Caption         =   "Your Version"
      Height          =   1380
      Left            =   180
      TabIndex        =   0
      Top             =   405
      Width           =   5670
      Begin VB.Label lblYourDate 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Released:"
         Height          =   195
         Left            =   645
         TabIndex        =   5
         Top             =   915
         Width           =   720
      End
      Begin VB.Label lblCurDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "pP---------------"
         Height          =   195
         Left            =   1410
         TabIndex        =   6
         Top             =   915
         Width           =   1080
      End
      Begin VB.Label lblYourEngine 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Engine:"
         Height          =   195
         Left            =   825
         TabIndex        =   3
         Top             =   585
         Width           =   540
      End
      Begin VB.Label lblCurEngVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00.000"
         Height          =   195
         Left            =   1410
         TabIndex        =   4
         Top             =   585
         Width           =   660
      End
      Begin VB.Label lblYourApp 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Application:"
         Height          =   195
         Left            =   525
         TabIndex        =   1
         Top             =   315
         Width           =   840
      End
      Begin VB.Label lblCurVer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0.00.000"
         Height          =   195
         Left            =   1410
         TabIndex        =   2
         Top             =   315
         Width           =   660
      End
   End
   Begin MSComctlLib.TabStrip tsVersions 
      Height          =   1875
      Left            =   45
      TabIndex        =   20
      Top             =   75
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   3307
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Your Version"
            Key             =   "tsYourVersion"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Latest Version"
            Key             =   "tsLatestVersion"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser wbWN 
      Height          =   2595
      Left            =   45
      TabIndex        =   14
      Top             =   2010
      Width           =   5895
      ExtentX         =   10398
      ExtentY         =   4577
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
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   45
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   5490
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   53
   End
   Begin InetCtlsObjects.Inet iCtrl 
      Left            =   300
      Top             =   5295
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSComctlLib.ProgressBar pbDl 
      Height          =   255
      Left            =   45
      TabIndex        =   16
      Top             =   4950
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4875
      TabIndex        =   18
      Top             =   5610
      Width           =   1065
   End
   Begin SHDocVwCtl.WebBrowser wbCtrl 
      Height          =   90
      Left            =   5235
      TabIndex        =   19
      Top             =   5775
      Width           =   75
      ExtentX         =   132
      ExtentY         =   159
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
      Caption         =   "pP---------------"
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
      Left            =   45
      TabIndex        =   15
      Top             =   4695
      Width           =   5895
   End
End
Attribute VB_Name = "frmUpgrade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cStep As Integer
Dim nvBuffer() As Byte
Dim StartDownloading As Boolean
Dim fSize As Long
Dim IsCancelling As Boolean
'Dim wskResponse As String

Private Sub cmdCancel_Click()

    If cStep = 2 Then
        If MsgBox("Are you sure you want to cancel the download?", vbYesNo + vbQuestion, "Confirm Cancel") = vbNo Then
            Exit Sub
        Else
            IsCancelling = True
            iCtrl.Cancel
        End If
    End If
    Unload Me

End Sub

Private Sub DoCheck()
    Dim ff As Integer
    Dim dlBytes As Long
    
    On Error GoTo DoCheck_Error

    If IsConnectedToInternet Then
        MousePointer = vbHourglass
        Me.Enabled = False
        
        DoEvents
    
        cStep = cStep + 1
    
        Select Case cStep
            Case 1
                cmdCheck.Enabled = False
                lblInfo.caption = GetLocalizedStr(464)
                wbCtrl.Navigate "https://xfx.net/utilities/dmbuilder/udl.php?noImgs=1"
            Case 2
                Dim p As Integer
                
                On Error Resume Next
                Kill TempPath + dlFileName
                On Error GoTo DoCheck_Error
                cmdCheck.Enabled = False
                cmdCancel.caption = GetLocalizedStr(187)
                lblInfo.caption = GetLocalizedStr(466)
                pbDl.Visible = True
                Me.Enabled = True
                
                ff = FreeFile
                Open TempPath + dlFileName For Binary Access Write As #ff
                    With iCtrl
                        .AccessType = icUseDefault
                        .protocol = icHTTPS
                        '.UserName = "xfxftp"
                        '.Password = "Tuu%1a09"
                        .Proxy = ""
                        .Execute "https://xfx.net/ftp/" + dlFileName
                        Do Until StartDownloading And Not .StillExecuting
                            DoEvents
                        Loop
                        nvBuffer() = iCtrl.GetChunk(8192, icByteArray)
                        Do
                            dlBytes = dlBytes + UBound(nvBuffer)
                            lblInfo.caption = GetLocalizedStr(467) + " " + dlFileName + " (" + NiceBytes(dlBytes) + " / " + NiceBytes(fSize) + ")"
                            p = (dlBytes / fSize) * 100
                            pbDl.Value = IIf(p > 100, 100, p)
                            Put #ff, , nvBuffer()
                            DoEvents
                            nvBuffer() = iCtrl.GetChunk(8192, icByteArray)
                        Loop Until UBound(nvBuffer) <= 0 Or IsCancelling
                        pbDl.Value = 100
                    End With
                Close #ff
                
                MousePointer = vbDefault
                cmdCheck.Enabled = True
                cmdCheck.caption = GetLocalizedStr(468)
            Case 3
                Going2Upgrade = True
                Unload Me
        End Select
    End If

   On Error GoTo 0
   Exit Sub

DoCheck_Error:
    MousePointer = vbDefault
    MsgBox "Error " & Err.number & " (" & Err.Description & ") in procedure DoCheck of Form frmUpgrade at line " & Erl
    
End Sub

Private Sub cmdCheck_Click()

    DoCheck

End Sub

Private Sub Form_Load()

    On Error Resume Next

    CenterForm Me
    SetupCharset Me
    DisableCloseButton Me
    LocalizeUI
    
    With frameYourVer
        frameLatestVer.Move .Left, .Top, .Width, .Height
        .ZOrder 0
    End With
    
    wbWN.Navigate "about:blank"
    
    cStep = 0

    lblRelDay.caption = ""
    lblInfo.caption = ""
    lblEngVer.caption = ""
    lblLatestVer.caption = ""
    lblCurVer.caption = DMBVersion
    lblCurEngVer.caption = Replace(EngineVersion, ".", GetDecimalSeparator)
    lblCurDate.caption = Format(CurEXEDate, "Long Date")
    lblSize.caption = ""
    
    TimerDoCheck.Enabled = True

End Sub

Private Sub Form_Resize()

    On Error Resume Next
    
    Dim ct As Long

    If WindowState = vbMinimized Then Exit Sub
    
    ct = GetClientTop(Me.hwnd)

    tsVersions.Width = Width - 195
    frameYourVer.Width = Width - 420
    frameLatestVer.Width = Width - 420
    
    wbWN.Move 45, 2010, Width - 195, Height - (3450 + ct)
    
    lblInfo.Move 45, wbWN.Top + wbWN.Height + 120, Width - 195, 210
    pbDl.Move 45, lblInfo.Top + lblInfo.Height + 30, Width - 195, 255
    uc3DLine1.Move 45, Height - (555 + ct), Width - 195, 30
    
    cmdCheck.Move Width - 2385, Height - (435 + ct)
    cmdCancel.Move Width - 1215, Height - (435 + ct)
    
    wbCtrl.Move Width \ 2, Height \ 2

End Sub

Private Sub iCtrl_StateChanged(ByVal State As Integer)

    Select Case State
        Case icResponseReceived
            StartDownloading = True
        Case icError
            MsgBox iCtrl.ResponseCode & ": " & iCtrl.ResponseInfo
    End Select

End Sub

Private Sub TimerDoCheck_Timer()

    TimerDoCheck.Enabled = False
    DoCheck

End Sub

Private Sub tsVersions_Click()

    Select Case tsVersions.SelectedItem.key
        Case "tsYourVersion"
            frameYourVer.ZOrder 0
        Case "tsLatestVersion"
            frameLatestVer.ZOrder 0
    End Select

End Sub

Private Sub wbCtrl_DocumentComplete(ByVal pDisp As Object, url As Variant)

    Dim doc As IHTMLDocument2
    Dim divs As IHTMLElementCollection
    Dim lVer As Single
    Dim lBuild As Integer
    Dim lEngineVer As Single
    Dim lEngineBuild As Integer
    Dim lDateDay As Integer
    Dim lDateMonth As Integer
    Dim sDateMonth As String
    Dim lDateYear As Integer
    Dim NeedUpgrade As Boolean
    Dim DateSep As String
    Dim DecSep As String
    
    Dim eVer As String
    
    If InStr(1, url, "whatsnew_info.htm") Then Exit Sub
    
    'On Error GoTo CheckFailed
    On Error Resume Next
    
    If cStep = 0 Then Exit Sub
    Do While wbCtrl.Document Is Nothing
        DoEvents
    Loop
    
    DateSep = GetDateSeparator
    DecSep = GetDecimalSeparator
    
    Set doc = wbCtrl.Document
    Set divs = doc.All.tags("span")
    
    cmdCheck.caption = GetLocalizedStr(469)
    cmdCheck.Enabled = True
    
    lVer = CSng(Replace(divs("Version3").innerText, ".", DecSep))
    lBuild = Val(divs("Build3").innerText)
    eVer = Replace(divs("Engine3").innerText, ".", DecSep)
    
    fSize = Val(divs("Size3").innerText)
    
    lEngineVer = CSng(Split(eVer, DecSep)(0) + DecSep + Split(eVer, DecSep)(1))
    lEngineBuild = Val(Split(eVer, DecSep)(2))
    
    If GetSetting("DMB", "RegInfo", "PreRegVer", 0) <> 0 Then
        dlFileName = GetFileName(divs("dlURL6").All.tags("a")(0))
    Else
        #If LITE = 1 Then
            dlFileName = GetFileName(divs("dlURL5").All.tags("a")(0))
        #Else
            #If DEVVER = 1 Then
                dlFileName = GetFileName(divs("dlURL4").All.tags("a")(0))
            #Else
                dlFileName = GetFileName(divs("dlURL3").All.tags("a")(0))
            #End If
        #End If
    End If
    
    If ResellerID <> "" Then dlFileName = Replace(dlFileName, ".exe", ResellerID + ".exe")
    
    lblLatestVer.caption = Format(Round(lVer, 2), "0.00") & DecSep & Format(lBuild, "000")
    lblEngVer.caption = Format(Round(lEngineVer, 2), "0.00") & DecSep & Format(lEngineBuild, "000")
    
    If lVer > CSng(App.Major & DecSep & App.Minor) Then NeedUpgrade = True
    If (lVer >= CSng(App.Major & DecSep & App.Minor)) And (lBuild > App.Revision) Then NeedUpgrade = True
    
    If lEngineVer > CSng((Split(EngineVersion, ".")(0) + DecSep + Split(EngineVersion, ".")(1))) Then NeedUpgrade = True
    If (lEngineVer >= CSng((Split(EngineVersion, ".")(0) + DecSep + Split(EngineVersion, ".")(1))) And (lEngineBuild > Val(Split(EngineVersion, ".")(2)))) Then NeedUpgrade = True
        
    lDateDay = Val(divs("DateDay").innerText)
    sDateMonth = divs("DateMonth").innerText
    lDateMonth = Switch(sDateMonth = "Jan", 1, sDateMonth = "Feb", 2, sDateMonth = "Mar", 3, sDateMonth = "Apr", 4, sDateMonth = "May", 5, sDateMonth = "Jun", 6, sDateMonth = "Jul", 7, sDateMonth = "Aug", 8, sDateMonth = "Sep", 9, sDateMonth = "Oct", 10, sDateMonth = "Nov", 11, sDateMonth = "Dec", 12)
    lDateYear = Val(divs("DateYear").innerText)
    Select Case Val(Left(GetLongDateFormat(), 1))
        Case 0
            lblRelDay.caption = Format(lDateMonth & DateSep & lDateDay & DateSep & lDateYear, "Long Date")
        Case 1
            lblRelDay.caption = Format(lDateDay & DateSep & lDateMonth & DateSep & lDateYear, "Long Date")
        Case 2
            lblRelDay.caption = Format(lDateYear & DateSep & lDateMonth & DateSep & lDateDay, "Long Date")
    End Select
    
    lblSize.caption = Format(fSize / 1048576, "0.00") + " MB"
    
    'If lDateYear > Year(Date) Then
    '    NeedUpgrade = True
    'ElseIf (lDateYear = Year(Date)) And (lDateMonth > Month(Date)) Then
    '    NeedUpgrade = True
    'ElseIf (lDateYear = Year(Date)) And (lDateMonth = Month(Date)) And (lDateDay > Day(Date)) Then
    '    NeedUpgrade = True
    'End If
    
    wbWN.Navigate "https://xfx.net/utilities/dmbuilder/whatsnew_info.htm"
    
    While wbWN.Document Is Nothing
        DoEvents
    Wend
    
    With wbWN.Document.body
        .topMargin = 0
        .leftMargin = 0
    End With
    
    'FORCED
    'NeedUpgrade = True
    
    'If chkNotify.Value = vbChecked Then SendEmail
    'chkNotify.Visible = False
    
    If NeedUpgrade Then
        lblInfo.caption = GetLocalizedStr(470)
    Else
        lblInfo.caption = GetLocalizedStr(471)
    End If
    
    cmdCheck.Visible = NeedUpgrade
    cmdCancel.caption = GetLocalizedStr(424)
    
    tsVersions.Tabs("tsLatestVersion").Selected = True
    tsVersions_Click
    
ExitSub:
    Me.Enabled = True
    MousePointer = vbDefault
    
    Exit Sub
    
CheckFailed:
    lblInfo.caption = GetLocalizedStr(472)
    lblRelDay.caption = ""
    lblInfo.caption = GetLocalizedStr(473)
    
    MsgBox GetLocalizedStr(474) + vbCrLf + "Error " & Err.number & ": " + Err.Description, vbCritical + vbOKOnly, GetLocalizedStr(663)
    
    GoTo ExitSub

End Sub

'Private Sub wskCtrl_DataArrival(ByVal bytesTotal As Long)
'
'    On Error Resume Next
'    wskCtrl.GetData wskResponse
'
'End Sub
'
'Private Function WaitFor(ResponseCode As String) As Boolean
'
'    Dim tmr As Long
'
'    tmr = Timer
'    While Len(wskResponse) = 0
'        DoEvents
'        If Timer - tmr > 10 Then Exit Function
'    Wend
'
'    tmr = Timer
'    While Left(wskResponse, 3) <> ResponseCode
'        DoEvents
'        If Timer - tmr > 10 Then Exit Function
'    Wend
'
'    WaitFor = True
'
'End Function
'
'Private Function Talk2Client(wCode As String, sData As String) As Boolean
'
'    On Error Resume Next
'
'    If WaitFor(wCode) Then
'        wskCtrl.SendData sData + vbCrLf
'        Talk2Client = True
'    Else
'        Talk2Client = False
'    End If
'
'    DoEvents
'
'End Function
'
'Private Sub SendEmail()
'
'    Dim EmailBody As String
'
'    On Error Resume Next
'
'    EmailBody = "From: DMB User" + vbCrLf + _
'                "To: Admin" + vbCrLf + _
'                "Subject: Upgrade Request" + vbCrLf + vbCrLf + _
'                "User Name: " + USER + vbCrLf + _
'                "Unlock Code: " + GetSetting(App.EXEName, "RegInfo", "SerialNumber", "DEMO")
'
'    If MsgBox("This information will be sent to the server: " + vbCrLf + vbCrLf + EmailBody + vbCrLf + vbCrLf + "Would you like to continue?", vbQuestion + vbYesNo, "Server Notification") = vbYes Then
'        wskCtrl.protocol = sckTCPProtocol
'        wskCtrl.RemoteHost = "xfx.net"
'        wskCtrl.RemotePort = 25
'        wskCtrl.Connect
'
'        While wskCtrl.State <> sckConnected
'            DoEvents
'        Wend
'
'        If Talk2Client("220", "HELO xfx.net") Then
'            If Talk2Client("250", "mail from: unk@unk.com") Then
'                If Talk2Client("250", "rcpt to: upgrade@xfx.net") Then
'                    If Talk2Client("250", "data") Then
'                        If Talk2Client("354", EmailBody + vbCrLf + ".") Then
'                            If Talk2Client("250", "quit") Then
'                                WaitFor "221"
'                            End If
'                        End If
'                    End If
'                End If
'            End If
'        End If
'
'        DoEvents
'        wskCtrl.Close
'    End If
'
'End Sub

Private Sub LocalizeUI()

    caption = GetLocalizedStr(462)

    tsVersions.Tabs("tsYourVersion").caption = GetLocalizedStr(456)
    tsVersions.Tabs("tsLatestVersion").caption = GetLocalizedStr(461)
    
    lblYourApp.caption = GetLocalizedStr(457) + ":"
    lblYourEngine.caption = GetLocalizedStr(458) + ":"
    lblYourDate.caption = GetLocalizedStr(459) + ":"
    
    lblLatestVerC.caption = GetLocalizedStr(457) + ":"
    lblLatestVerE.caption = GetLocalizedStr(458) + ":"
    lblReleased.caption = GetLocalizedStr(459) + ":"
    
    lblInfo.caption = GetLocalizedStr(463)
    
    cmdCheck.caption = GetLocalizedStr(460)
    cmdCancel.caption = GetLocalizedStr(187)
    
    If Preferences.language <> "eng" Then
        cmdCheck.Width = SetCtrlWidth(cmdCheck)
        cmdCancel.Width = SetCtrlWidth(cmdCancel)
    End If

End Sub
