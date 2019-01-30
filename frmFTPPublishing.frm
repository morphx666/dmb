VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Begin VB.Form frmFTPPublishing 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "FTP Publishing"
   ClientHeight    =   2505
   ClientLeft      =   6825
   ClientTop       =   6120
   ClientWidth     =   6105
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFTPPublishing.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtDummy 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   2625
      TabIndex        =   10
      Top             =   2190
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.CommandButton cmdViewLog 
      Caption         =   "View Log..."
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
      Left            =   75
      TabIndex        =   9
      Top             =   2055
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Timer tmrInit 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3660
      Top             =   1140
   End
   Begin VB.CheckBox chkImages 
      Caption         =   "Publish Images"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1110
      TabIndex        =   1
      Top             =   1545
      Value           =   1  'Checked
      Width           =   4245
   End
   Begin VB.TextBox txtInfo 
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
      ForeColor       =   &H80000011&
      Height          =   390
      Left            =   1110
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "frmFTPPublishing.frx":27A2
      Top             =   690
      Width           =   4875
   End
   Begin MSComctlLib.ProgressBar pbProgress 
      Height          =   270
      Left            =   1110
      TabIndex        =   7
      Top             =   360
      Width           =   4350
      _ExtentX        =   7673
      _ExtentY        =   476
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdPublish 
      Caption         =   "OK"
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
      Left            =   4005
      TabIndex        =   2
      Top             =   2055
      Width           =   900
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
      Left            =   5130
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Top             =   2055
      Width           =   900
   End
   Begin VB.CheckBox chkJSFiles 
      Caption         =   "Publish Menus"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   1110
      TabIndex        =   0
      Top             =   1290
      Value           =   1  'Checked
      Width           =   4245
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   930
      TabIndex        =   5
      Top             =   1140
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   53
   End
   Begin InetCtlsObjects.Inet inetFTP 
      Left            =   210
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin xfxLine3D.ucLine3D uc3DLine2 
      Height          =   30
      Left            =   15
      TabIndex        =   6
      Top             =   1905
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   53
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   180
      Picture         =   "frmFTPPublishing.frx":27B0
      Top             =   270
      Width           =   480
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "FTP Publishing Progress"
      Height          =   210
      Left            =   1110
      TabIndex        =   4
      Top             =   105
      Width           =   1920
   End
End
Attribute VB_Name = "frmFTPPublishing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private ThisConfig As ConfigDef
Private TotalBytes As Long
Private UppedBytes As Long
Private UserCancel As Boolean
Private Images() As String
Private LastPaths As String
Private log As String

Private ftpHost As String
Private ftpPath As String
Private ftpProxyServer As String
Private ftpProxyPort As String

Private logFile As String

Private doAbort As Boolean

Private Sub chkImages_Click()

    cmdPublish.Enabled = (chkImages.Value = vbChecked) Or (chkJSFiles.Value = vbChecked)

End Sub

Private Sub chkJSFiles_Click()

    cmdPublish.Enabled = (chkImages.Value = vbChecked) Or (chkJSFiles.Value = vbChecked)

End Sub

Private Sub cmdCancel_Click()

    If cmdPublish.Enabled = False Then
        UserCancel = (MsgBox("Are you sure you want to cancel Publishing?", vbQuestion + vbYesNo, "Publishing") = vbYes)
        Exit Sub
    End If
    
    Unload Me

End Sub

Private Sub cmdPublish_Click()
    
    UserCancel = False

    cmdPublish.Enabled = False
    MousePointer = vbHourglass
    LastPaths = ""
    chkImages.Enabled = False
    chkJSFiles.Enabled = False
    cmdViewLog.Visible = False
    
    Me.Enabled = False
    If Project.HasChanged Then
        txtInfo.Text = "Compiling..."
        frmMain.ToolsCompile
    End If
    Me.Enabled = True
    StartFTPTX
    
    'If Not doAbort Then
        On Error Resume Next
        inetFTP.Cancel
        
        cmdPublish.Enabled = True
        cmdPublish.Visible = False
        cmdCancel.Enabled = True
        cmdCancel.Caption = "Close"
        cmdViewLog.Visible = True
        MousePointer = vbDefault
        
        DoEvents
        
        logFile = GetFileName(Project.fileName)
        logFile = Left(logFile, Len(logFile) - 4)
        logFile = GetFilePath(Project.fileName) + logFile + ".log"
        
        ff = FreeFile
        Open logFile For Output As #ff
            Print #ff, "Transfer Log for '" + Project.Name + "'"
            Print #ff, Format(Date, "Long Date") + " " + Format(Time, "Long Time")
            Print #ff, String(35, "*") + vbCrLf
            Print #ff, log
        Close #ff
    'End If
    
End Sub

Private Sub cmdViewLog_Click()

    On Error Resume Next
    
    Err.Clear
    Shell "notepad.exe " + logFile, vbNormalFocus
    
    If Err.number <> 0 Then
        MsgBox Err.number & ": " & Err.Description, vbCritical + vbOKOnly, "Unable to display log file"
    End If

End Sub

Private Sub Form_Load()

    CenterForm Me
    
    log = ""
    UppedBytes = 0
    txtInfo.Text = ""
    
    ftpUserName = ""
    ftpPassword = ""
    doAbort = False
    
    tmrInit.Enabled = True

End Sub

Private Sub StartFTPTX()

    Dim JSPath As String
    Dim imgPath As String
    Dim i As Integer
    
    JSPath = ThisConfig.CompiledPath
    
    With inetFTP
        If ftpProxyServer <> "" Then
            .AccessType = icNamedProxy
            .Proxy = ftpProxyServer + ":" + ftpProxyPort
        Else
            .AccessType = icDirect
        End If
        .protocol = icFTP
        .RequestTimeout = 30
        
        .RemoteHost = Replace(Replace(ftpHost, "ftp://", ""), "/", "")
        If ftpUserName = "anonymous" Then
            .UserName = ftpUserName
            .Password = "dmbuser@"
        Else
            If ftpUserName = "" Or ftpPassword = "" Then
                With frmFTPAccountInfo
                    .txtUserName = ftpUserName
                    .txtPassword = ftpPassword
                    
                    .Show vbModal
                End With
            End If
            
            If ftpUserName = "" Or ftpPassword = "" Then
                ErrorOut
                Exit Sub
            Else
                .UserName = ftpUserName
                .Password = ftpPassword
            End If
        End If
        
        pbProgress.Max = 100
        CopyProjectImages TempPath
        If chkImages.Value = vbChecked Then
            TotalBytes = CalcImagesSize(Images)
        Else
            TotalBytes = 0
        End If
        
        If DoLogin Then
            If chkJSFiles.Value = vbChecked Then
                Select Case Project.CodeOptimization
                    Case cocDEBUG
                        UploadJSFile JSPath + Project.JSFileName + ".js"
                        UploadJSFile JSPath + Project.JSFileName + "_frames.js"
                    Case Else
                        UploadJSFile JSPath + "ie" + Project.JSFileName + ".js"
                        UploadJSFile JSPath + "ie" + Project.JSFileName + "_frames.js"
                        UploadJSFile JSPath + "ns" + Project.JSFileName + ".js"
                        UploadJSFile JSPath + "ns" + Project.JSFileName + "_frames.js"
                End Select
            End If
            
            If chkImages.Value = vbChecked Then
                For i = 1 To UBound(Images)
                    Upload TempPath + GetFileName(Images(i))
                Next i
            End If
        Else
            MsgBox "Error " & .ResponseCode & ": " & .ResponseInfo, vbCritical + vbOKOnly, "Publishing Error"
        End If
    End With
    
    txtInfo.Text = "Publishing Complete"

End Sub

Private Sub UploadJSFile(fileName As String)

    If FileExists(fileName) Then
        TotalBytes = TotalBytes + FileLen(fileName)
        Upload fileName
    End If

End Sub

Private Function DoLogin() As Boolean

    On Error Resume Next

    txtInfo.Text = "Logging in..."
    DoEvents

    inetFTP.Execute , "CD ."
    Do
        DoEvents
    Loop While inetFTP.StillExecuting And Not UserCancel
    
    If inetFTP.ResponseCode = 0 Then
        DoLogin = True
    Else
        DoLogin = False
    End If

End Function

Private Sub Upload(srcFile As String)

    Dim tURL As String
    Dim tFile As String
    Dim tPath As String
    Dim paths() As String
    Dim mPath As String
    Dim i As Integer
    
    On Error GoTo ReportError
    
    If Not FileExists(srcFile) Then Exit Sub
    If UserCancel Then Exit Sub
    
    With inetFTP
        tFile = GetFileName(srcFile)

        If Right(srcFile, 2) = "js" Then
            tPath = FTP_JSPath
        Else
            tPath = FTP_IMGPath
        End If
        tURL = Replace("/" + tPath + tFile, "//", "/")
        
        If tPath <> "/" And tPath <> LastPaths Then
            LastPaths = tPath
            paths = Split(tPath, "/")
            For i = 1 To UBound(paths) - 1
                mPath = mPath + "/" + paths(i)
                
                txtInfo.Text = "Creating Path " + .RemoteHost + mPath
                log = log + "MKDIR " + mPath + vbCrLf
                .Execute , "MKDIR " + mPath
                Do
                    DoEvents
                Loop While .StillExecuting And Not UserCancel
                If UserCancel Then Exit Sub
            Next i
        End If
        
        txtInfo.Text = "Uploading " + GetFileName(srcFile) + " to" + vbCrLf + tURL
        log = log + Replace(txtInfo.Text, vbCrLf, " ") + vbCrLf
        .Execute , "CD " + SetSlashDir(GetFilePath(tURL, True), sdFwd)
        Do
            DoEvents
        Loop While .StillExecuting And Not UserCancel
        .Execute , "DELETE " + GetFileName(srcFile)
        Do
            DoEvents
        Loop While .StillExecuting And Not UserCancel
        
        .Execute , "PUT " + Long2Short(srcFile) + " " + tURL
        Do
            DoEvents
        Loop While .StillExecuting And Not UserCancel
        
        .Execute , "PUT " + Long2Short(srcFile) + " " + tURL
        Do
            DoEvents
        Loop While .StillExecuting And Not UserCancel
    End With
    
    UppedBytes = UppedBytes + FileLen(srcFile)
    
    On Error Resume Next
    pbProgress.Value = UppedBytes / TotalBytes * 100
    
    Exit Sub
    
ReportError:
    DoEvents
    MsgBox "An error has accoured in the FTP process." + vbCrLf + vbCrLf + _
           "Error " & Err.number & ": " & Err.Description + _
           IIf(inetFTP.ResponseCode <> 0, vbCrLf + "FTP Response " & inetFTP.ResponseCode & ": " & inetFTP.ResponseInfo, "") _
           , vbInformation + vbOKOnly, "FTP Error"
    UserCancel = True
    doAbort = True

End Sub

Private Sub inetFTP_StateChanged(ByVal State As Integer)

    If inetFTP.ResponseCode <> 0 Then
        Select Case State
            Case icResolvingHost:       log = log + "Resolving Host" + vbCrLf
            Case icHostResolved:        log = log + "Host Resolved" + vbCrLf
            Case icConnecting:          log = log + "Connecting" + vbCrLf
            Case icConnected:           log = log + "Connected" + vbCrLf
            Case icRequesting:          log = log + "Requesting" + vbCrLf
            Case icRequestSent:         log = log + "Request Sent" + vbCrLf
            Case icReceivingResponse:   log = log + "Receiving Response" + vbCrLf
            Case icResponseReceived:    log = log + "Response Received" + " (" & inetFTP.ResponseCode & "): " + inetFTP.ResponseInfo + vbCrLf
            Case icDisconnecting:       log = log + "Disconnecting" + vbCrLf
            Case icDisconnected:        log = log + "Disconnected" + vbCrLf
            Case icResponseCompleted:   log = log + "Response Completed" + vbCrLf
            Case icError:               log = log + "Error (" & inetFTP.ResponseCode & "): " + inetFTP.ResponseInfo + vbCrLf
        End Select
    End If
    
End Sub

Private Sub tmrInit_Timer()

    tmrInit.Enabled = False

    ThisConfig = Project.UserConfigs(Project.DefaultConfig)
    If Project.UserConfigs(Project.DefaultConfig).Type = ctcRemote Then ThisConfig = GetRealLocal
    
    Dim ftpInfo() As String
    If ThisConfig.FTP = "" Then
        ReDim ftpInfo(6)
    Else
        ftpInfo = Split(ThisConfig.FTP, "*")
    End If
    If UBound(ftpInfo) >= 0 Then ftpHost = ftpInfo(0)
    If UBound(ftpInfo) >= 1 Then ftpPath = ftpInfo(1)
    If UBound(ftpInfo) >= 2 Then ftpUserName = ftpInfo(2)
    If UBound(ftpInfo) >= 3 Then ftpPassword = ftpInfo(3)
    If UBound(ftpInfo) >= 4 Then
        If UBound(ftpInfo) >= 5 Then ftpProxyServer = ftpInfo(5)
        If UBound(ftpInfo) >= 6 Then ftpProxyPort = ftpInfo(6)
    End If
    
    If ftpHost = "" Or ftpPath = "" Then
        ErrorOut
    Else
        If Right(ftpHost, 1) = "/" Then ftpHost = Left(ftpHost, Len(ftpHost) - 1)
        txtDummy.Width = txtInfo.Width
        txtInfo.Text = "Publish '" + Project.Name + "' menu to:" + vbCrLf + SetSlashDir(EllipseText(txtDummy, SetSlashDir(ftpHost + FTP_JSPath, sdBack), DT_PATH_ELLIPSIS), sdFwd)
    End If

End Sub

Private Function FTP_JSPath() As String

    On Error Resume Next
    FTP_JSPath = Replace(ftpPath + SetSlashDir(Mid(ThisConfig.CompiledPath, Len(ThisConfig.RootWeb)), sdFwd), "//", "/")

End Function

Private Function FTP_IMGPath() As String

    On Error Resume Next
    FTP_IMGPath = Replace(ftpPath + SetSlashDir(Mid(ThisConfig.ImagesPath, Len(ThisConfig.RootWeb)), sdFwd), "//", "/")

End Function

Private Sub ErrorOut()
    doAbort = True
    
    MsgBox "FTP Information is missing or incomplete" + vbCrLf + "Click File->Project Properties->Configurations, select the default configuration and then click on the 'FTP Information' button to edit the Publishing parameters.", vbCritical + vbOKOnly, "Unable to Publish Menus"
    Unload Me
End Sub
