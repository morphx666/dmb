VERSION 5.00
Object = "{9A2947B4-87EE-40EE-A3EF-32BDC32D5726}#1.0#0"; "xfxline3d.ocx"
Object = "{2A2AD7CA-AC77-46F3-84DC-115021432312}#1.0#0"; "hRef.ocx"
Begin VB.Form frmCompilationReport 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compilation Report"
   ClientHeight    =   5655
   ClientLeft      =   6270
   ClientTop       =   4890
   ClientWidth     =   5625
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCompilationReport.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkAutoShowReport 
      Caption         =   "Display After Compiling"
      Height          =   195
      Left            =   180
      TabIndex        =   22
      Top             =   5310
      Width           =   2220
   End
   Begin VB.ComboBox cmbMode 
      Height          =   315
      ItemData        =   "frmCompilationReport.frx":0E42
      Left            =   375
      List            =   "frmCompilationReport.frx":0E4C
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   4155
      Width           =   1830
   End
   Begin VB.ComboBox cmbEDT 
      Height          =   315
      ItemData        =   "frmCompilationReport.frx":0E6B
      Left            =   375
      List            =   "frmCompilationReport.frx":0E81
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   4575
      Width           =   1830
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "Close"
      Default         =   -1  'True
      Height          =   375
      Left            =   4665
      TabIndex        =   21
      Top             =   5220
      Width           =   900
   End
   Begin xfxLine3D.ucLine3D uc3DLine3 
      Height          =   30
      Left            =   2430
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   3930
      Width           =   2805
      _ExtentX        =   4948
      _ExtentY        =   53
   End
   Begin xfxLine3D.ucLine3D uc3DLine1 
      Height          =   30
      Left            =   930
      TabIndex        =   1
      Top             =   285
      Width           =   4305
      _ExtentX        =   7594
      _ExtentY        =   53
   End
   Begin xfxLine3D.ucLine3D uc3DLine2 
      Height          =   30
      Left            =   1200
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1425
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   53
   End
   Begin xfxLine3D.ucLine3D uc3DLine4 
      Height          =   30
      Left            =   45
      TabIndex        =   20
      Top             =   5085
      Width           =   5565
      _ExtentX        =   9816
      _ExtentY        =   53
   End
   Begin href.uchref1 ucViewMenus 
      Height          =   315
      Left            =   2160
      TabIndex        =   23
      Top             =   1020
      Width           =   1650
      _ExtentX        =   2910
      _ExtentY        =   556
      Caption         =   "View Compiled Menus"
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
   End
   Begin VB.Label lblEDT 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00h : 00m : 00s"
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2610
      TabIndex        =   18
      Top             =   4380
      Width           =   2250
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Estimated Download Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   15
      Top             =   3825
      Width           =   2145
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "Monotype.com"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Index           =   0
      Left            =   2205
      TabIndex        =   3
      Top             =   442
      Visible         =   0   'False
      Width           =   720
   End
   Begin VB.Line Line1 
      X1              =   240
      X2              =   5205
      Y1              =   3210
      Y2              =   3210
   End
   Begin VB.Label lblTotal 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
      Height          =   195
      Left            =   375
      TabIndex        =   14
      Top             =   3270
      Width           =   360
   End
   Begin VB.Label lblImages 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Images and Sounds"
      Height          =   195
      Left            =   375
      TabIndex        =   13
      Top             =   2940
      Width           =   1410
   End
   Begin VB.Label lblOpNSFrames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "frames_nsmenu.js"
      Height          =   195
      Left            =   375
      TabIndex        =   12
      Top             =   2715
      Width           =   1320
   End
   Begin VB.Label lblOpIEFrames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "frames_iemenu.js"
      Height          =   195
      Left            =   375
      TabIndex        =   11
      Top             =   2490
      Width           =   1275
   End
   Begin VB.Label lblOpNS 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "nsmenu.js"
      Height          =   195
      Left            =   375
      TabIndex        =   10
      Top             =   2250
      Width           =   735
   End
   Begin VB.Label lblOpIE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "iemenu.js"
      Height          =   195
      Left            =   375
      TabIndex        =   9
      Top             =   2025
      Width           =   690
   End
   Begin VB.Label lblNoOpFrames 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "frames.js"
      Height          =   195
      Left            =   375
      TabIndex        =   8
      Top             =   1800
      Width           =   675
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Size Report"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   5
      Top             =   1305
      Width           =   930
   End
   Begin VB.Label lblNoOp 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "menu.js"
      Height          =   195
      Left            =   375
      TabIndex        =   7
      Top             =   1575
      Width           =   570
   End
   Begin VB.Label lblConf 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Configuration:"
      Height          =   195
      Left            =   375
      TabIndex        =   4
      Top             =   660
      Width           =   1035
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   210
      Left            =   180
      TabIndex        =   0
      Top             =   165
      Width           =   675
   End
   Begin VB.Label lblOpM 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Optimization Method:"
      Height          =   195
      Left            =   375
      TabIndex        =   2
      Top             =   450
      Width           =   1530
   End
End
Attribute VB_Name = "frmCompilationReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Total As Long
Dim ieTotal As Long
Dim nsTotal As Long
Dim imgTotal As Long

Private Enum rscTypeConstants
    rtcIENS = 0
    rtcIE = 1
    rtcNS = 2
    rtcIMG = 3
    rtcT = 4
End Enum

Private Sub cmbEDT_Click()

    UpdateEDT

End Sub

Private Sub UpdateEDT()
    
    Dim b As Long
    Dim h As Integer
    Dim m As Integer
    Dim s As Integer
    Dim bTotal As Long
    
    Select Case cmbMode.ListIndex
        Case 1 'Navigator 4
            bTotal = nsTotal
        Case 0, -1 'DOM Compliant browsers
            bTotal = ieTotal
    End Select

    Select Case cmbEDT.ListIndex
        Case 0, -1 'Modem (56kbps)
            b = 56000
        Case 1 'ISDN (128kbps)
            b = 128000
        Case 2 'DSL (256kbps)
            b = 256000
        Case 3 'T1 (1mbps)
            b = 1000000
        Case 4 'LAN (10mbps)
            b = 10000000
        Case 5 'LAN (100mbps)
            b = 100000000
    End Select
    
    b = (bTotal + imgTotal) * 8 / b
    
    h = Int(b / 36000)
    m = Int(b / 60) - h * 60
    s = b - ((h * 3600) + (m * 60))
    
    lblEDT.Caption = h & "h : " & m & "m : " & s & "s"

End Sub

Private Sub cmbMode_Click()

    UpdateEDT

End Sub

Private Sub cmdClose_Click()

    Preferences.AutoShowCompileReport = (chkAutoShowReport.Value = vbChecked)
    
    SaveSetting App.EXEName, "CompileReport", "BrowserSel", cmbMode.ListIndex
    SaveSetting App.EXEName, "CompileReport", "LineSpeed", cmbEDT.ListIndex
    
    Unload Me

End Sub

Private Sub Form_Load()

    lblImages.Caption = "Images"

    CenterForm Me
    InitDlg

End Sub

Private Sub InitDlg()

    Dim ThisConfig As ConfigDef
    Dim LocalConfig As ConfigDef
    Dim Images() As String
    Dim ImagesTotal As Long
    
    chkAutoShowReport.Value = IIf(Preferences.AutoShowCompileReport, vbChecked, vbUnchecked)
    
    Total = 0
    ieTotal = 0
    nsTotal = 0
    imgTotal = 0
    
    ThisConfig = Project.UserConfigs(Project.DefaultConfig)
    
    If ThisConfig.Type = ctcRemote Then
        LocalConfig = Project.UserConfigs(GetConfigID(ThisConfig.LocalInfo4RemoteConfig))
        ucViewMenus.url = ThisConfig.RootWeb
    Else
        LocalConfig = ThisConfig
        If LocalConfig.HotSpotEditor.HotSpotsFile <> "" Then
            Dim s As String
            s = QueryValue(HKEY_CLASSES_ROOT, "." + GetFileExtension(LocalConfig.HotSpotEditor.HotSpotsFile))
            If s <> "" Then
                s = QueryValue(HKEY_CLASSES_ROOT, s + "\shell\open\command")
                ucViewMenus.url = Replace(s, "%1", Long2Short(LocalConfig.HotSpotEditor.HotSpotsFile))
            Else
                ucViewMenus.Enabled = False
            End If
        Else
            ucViewMenus.url = Long2Short(LocalConfig.CompiledPath)
        End If
    End If
    
    lblNoOp.Caption = Project.JSFileName + ".js"
    lblNoOpFrames.Caption = "frames_" + Project.JSFileName + ".js"
    lblOpIE.Caption = "ie" + Project.JSFileName + ".js"
    lblOpNS.Caption = "ns" + Project.JSFileName + ".js"
    lblOpIEFrames = "frames_ie" + Project.JSFileName + ".js"
    lblOpNSFrames = "frames_ns" + Project.JSFileName + ".js"
    
    AddLabel lblConf, ThisConfig.Name
    With AddLabel(lblConf, LCase(GetRealLocal.CompiledPath))
        .FontSize = 7
        .FontBold = False
        .Top = .Top + lblConf.Height
    End With
    
    If Project.CodeOptimization = cocDEBUG Then
        AddLabel lblOpM, GetLocalizedStr(822)
        
        AddLabel lblNoOp, FormatSize(LocalConfig.CompiledPath + Project.JSFileName + ".js", 0, rtcIENS)
        AddLabel lblOpIE, "0.00 KB (not used)"
        AddLabel lblOpNS, "0.00 KB (not used)"
        
        If ThisConfig.Frames.UseFrames Then
            AddLabel lblNoOpFrames, FormatSize(LocalConfig.CompiledPath + "ie" + Project.JSFileName + ".js", 0, rtcIENS)
        Else
            AddLabel lblNoOpFrames, "0.00 KB (not used)"
            AddLabel lblOpIEFrames, "0.00 KB (not used)"
            AddLabel lblOpNSFrames, "0.00 KB (not used)"
        End If
    Else
        If Project.UseGZIP Then lblOpIE.Caption = "iemenu.js.gz"
        If Project.UseGZIP Then lblOpIEFrames.Caption = "frames_iemenu.js.gz"
        
        AddLabel lblOpM, IIf(Project.CodeOptimization = cocNormal, GetLocalizedStr(823), GetLocalizedStr(824))
        
        AddLabel lblNoOp, "0.00 KB (not used)"
        AddLabel lblNoOpFrames, "0.00 KB (not used)"
        AddLabel lblOpIE, FormatSize(LocalConfig.CompiledPath + "ie" + Project.JSFileName + ".js" + IIf(Project.UseGZIP, ".gz", ""), 0, rtcIE)
        AddLabel lblOpNS, FormatSize(LocalConfig.CompiledPath + "ns" + Project.JSFileName + ".js", 0, rtcNS)
        
        If ThisConfig.Frames.UseFrames Then
            AddLabel lblOpIEFrames, FormatSize(LocalConfig.CompiledPath + "ie" + Project.JSFileName + "_frames.js" + IIf(Project.UseGZIP, ".gz", ""), 0, rtcIE)
            AddLabel lblOpNSFrames, FormatSize(LocalConfig.CompiledPath + "ns" + Project.JSFileName + "_frames.js", 0, rtcNS)
        Else
            AddLabel lblOpIEFrames, "0.00 KB (not used)"
            AddLabel lblOpNSFrames, "0.00 KB (not used)"
        End If
    End If
    
    CopyProjectImages TempPath
    ImagesTotal = CalcImagesSize(Images)
    Total = Total + ImagesTotal
    AddLabel lblImages, FormatSize("", ImagesTotal, rtcIMG)
    AddLabel lblTotal, FormatSize("", Total, rtcT)
    
    cmbMode.ListIndex = GetSetting(App.EXEName, "CompileReport", "BrowserSel", 0)
    cmbEDT.ListIndex = GetSetting(App.EXEName, "CompileReport", "LineSpeed", 2)

End Sub

Private Function AddLabel(ref As Label, c As String) As Label

    Dim s() As String
    Static refPos As Integer
    Dim p As Integer
    
    p = InStr(c, " KB ")
    If p Then
        s = Split(c, " KB ")
    Else
        ReDim s(1)
        s(0) = c
        s(1) = ""
        refPos = 0
    End If

    Load lblInfo(lblInfo.Count)
    With lblInfo(lblInfo.Count - 1)
        .Caption = s(0)
        DoEvents
        .Move lblInfo(0).Left, ref.Top + ref.Height / 2 - .Height / 2
        If p Then
            .Caption = .Caption + " KB "
            .Width = GetTextSize(.Caption, , lblInfo(lblInfo.Count - 1))(1)
            If refPos = 0 Then
                refPos = (.Left + .Width)
            ElseIf lblInfo.Count > 1 Then
                .Left = refPos - .Width
            End If
            .Caption = .Caption + s(1)
        End If
        If InStr(c, "not used") Then
            .Enabled = False
            ref.Enabled = False
        End If
        .Visible = True
    End With
    
    Set AddLabel = lblInfo(lblInfo.Count - 1)

End Function

Private Function FormatSize(ByVal File As String, ByVal s As Long, tg As rscTypeConstants) As String

    Dim fs As String
    
    On Error GoTo FileNotFound
    
    If s = 0 Then
        s = FileLen(File)
        Total = Total + s
        
        If Project.CodeOptimization = cocDEBUG Then
            If GetFileName(File) = Project.JSFileName + ".js" Then
                ieTotal = ieTotal + s
                nsTotal = nsTotal + s
            Else
                imgTotal = imgTotal + s
            End If
        Else
            Select Case tg
                Case rtcIE: ieTotal = ieTotal + s
                Case rtcNS: nsTotal = nsTotal + s
                Case rtcIMG: imgTotal = imgTotal + s
            End Select
        End If
    Else
        imgTotal = imgTotal + s
    End If
    
    fs = Format(s / 1024, "0.00") + " KB (" + Format(s, "0,000.00") + " bytes)"
    If s / 1024 < 10 Then fs = " " + fs
    
    FormatSize = fs
    
    Exit Function
    
FileNotFound:
    FormatSize = " 0.00 KB (File Not Found)"
    
End Function
