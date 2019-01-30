VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmHotSpotEditor 
   Caption         =   "HotSpot Editor"
   ClientHeight    =   7350
   ClientLeft      =   4560
   ClientTop       =   4080
   ClientWidth     =   8985
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmHotSpotEditor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7350
   ScaleWidth      =   8985
   Begin VB.Timer tmrDoneLoading 
      Enabled         =   0   'False
      Interval        =   150
      Left            =   7185
      Top             =   3450
   End
   Begin VB.Timer tmrMonitorLoad 
      Interval        =   50
      Left            =   4095
      Top             =   6135
   End
   Begin VB.Timer tmrWait 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   2520
      Top             =   6120
   End
   Begin VB.Timer tmrSaveClose 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   3285
      Top             =   6015
   End
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1830
      Top             =   5880
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   330
      Left            =   6405
      TabIndex        =   6
      Top             =   6945
      Width           =   975
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "&Close"
      Height          =   330
      Left            =   7515
      TabIndex        =   5
      Top             =   6945
      Width           =   975
   End
   Begin VB.PictureBox cHotSpot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillColor       =   &H0000C000&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Index           =   0
      Left            =   555
      ScaleHeight     =   315
      ScaleWidth      =   510
      TabIndex        =   4
      Top             =   5955
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picMenuShape 
      Appearance      =   0  'Flat
      BackColor       =   &H00808000&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   1395
      ScaleHeight     =   525
      ScaleWidth      =   1245
      TabIndex        =   2
      Top             =   1215
      Visible         =   0   'False
      Width           =   1275
   End
   Begin VB.PictureBox picElement 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   135
      ScaleHeight     =   345
      ScaleWidth      =   540
      TabIndex        =   1
      Top             =   4800
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.Timer tmrMonitor 
      Enabled         =   0   'False
      Interval        =   75
      Left            =   210
      Top             =   4140
   End
   Begin SHDocVwCtl.WebBrowser wbCtrl 
      Height          =   3195
      Left            =   1050
      TabIndex        =   0
      Top             =   2100
      Width           =   4320
      ExtentX         =   7620
      ExtentY         =   5636
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
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
   Begin MSComctlLib.Toolbar tbOptions 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   635
      ButtonWidth     =   1323
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   3500
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   2700
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Next"
            Key             =   "tbNext"
            Object.ToolTipText     =   "Next Menu"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "tbHelp"
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageCombo icmbAlignment 
         Height          =   330
         Left            =   4515
         TabIndex        =   13
         Top             =   0
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   582
         _Version        =   393216
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Locked          =   -1  'True
      End
      Begin VB.PictureBox picAlignment 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   3615
         ScaleHeight     =   315
         ScaleWidth      =   2175
         TabIndex        =   11
         Top             =   15
         Width           =   2175
         Begin VB.Label lblAlignment 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Alignment"
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
            Left            =   165
            TabIndex        =   12
            Top             =   60
            Width           =   705
         End
      End
      Begin VB.PictureBox picGroups 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   30
         ScaleHeight     =   330
         ScaleWidth      =   3360
         TabIndex        =   8
         Top             =   0
         Width           =   3360
         Begin MSComctlLib.ImageCombo cmbMenuGroups 
            Height          =   330
            Left            =   555
            TabIndex        =   9
            Top             =   0
            Width           =   2790
            _ExtentX        =   4921
            _ExtentY        =   582
            _Version        =   393216
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Locked          =   -1  'True
         End
         Begin VB.Label lblGroups 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Groups"
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
            Left            =   15
            TabIndex        =   10
            Top             =   75
            Width           =   510
         End
      End
   End
   Begin SHDocVwCtl.WebBrowser wbCtrl2 
      Height          =   1740
      Left            =   5520
      TabIndex        =   7
      Top             =   4575
      Visible         =   0   'False
      Width           =   2715
      ExtentX         =   4789
      ExtentY         =   3069
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
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
End
Attribute VB_Name = "frmHotSpotEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If DEMO = 0 Then

Private Enum UserActionConstants
    [uaSAVE]
    [uaCLOSE]
End Enum
Dim UserAction As UserActionConstants

Dim doc As IHTMLDocument2
Dim DisableNav As Boolean
Dim mx As Long
Dim my As Long
Dim txtClickPos As Single
Dim ChDC As Long
Dim cEInfo As IHTMLElement
Dim ofLeft, ofTop As Long
Dim HotSpots() As IHTMLElement
Dim FoundAttachedMenus As Integer
Dim ConLost As Boolean
Dim ChangesSaved As Boolean
Dim MonitorBusy As Boolean
Dim IsBusy As Boolean

Dim pg1 As Integer
Dim pg2 As Integer
Dim defCaption As String

Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Sub cHotSpot_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = vbLeftButton Then
        picElement.Visible = False
        Refresh
        DoEvents
        HighlightCurrent
        picElement.Visible = True
    End If

End Sub

Private Sub cmbMenuGroups_Click()

    UpdateAlignmentCombo
    
    If Not doc Is Nothing Then
        Refresh
        DoEvents
        HighlightNewHotSpots
    End If

End Sub

Private Sub UpdateAlignmentCombo()

    Dim i As Integer

    For i = 1 To icmbAlignment.ComboItems.Count
        If Val(icmbAlignment.ComboItems(i).Tag) = MenuGrps(cmbMenuGroups.SelectedItem.Index).Alignment Then
            icmbAlignment.ComboItems(i).Selected = True
            Exit For
        End If
    Next i

End Sub

Private Sub cmdClose_Click()
    
    CleanUp
    UserAction = uaCLOSE
    tmrWait.Enabled = True

End Sub

Private Sub InitClose()

    HSCanceled = True
    tmrSaveClose.Enabled = True

End Sub

Private Sub cmdSave_Click()

    CleanUp
    UserAction = uaSAVE
    tmrWait.Enabled = True

End Sub

Private Sub InitSave()

    ChangesSaved = True
    SaveHotSpots
    HSCanceled = False
    Me.Visible = False

End Sub

Private Sub CleanUp()

    Dim ctrl As Control

    For Each ctrl In Controls
        DoEvents
        If TypeOf ctrl Is Timer Then
            ctrl.Enabled = False
        End If
    Next ctrl
    
    MousePointer = vbHourglass

End Sub

Private Function RemoveLocalPaths(ByVal cCode As String) As String

    Dim hfPath As String
    Dim i As Integer
    
    On Error GoTo ReturnDefault

    If InStr(cCode, "file:///") Then
        hfPath = Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile
        hfPath = Left(hfPath, InStrRev(hfPath, "\"))
        hfPath = LCase("file:///" + SetSlashDir(hfPath, sdFwd))
        
        For i = 32 To 255
            cCode = Replace(cCode, "%" + Format(Hex(i), "00"), Chr(i))
            hfPath = Replace(hfPath, "%" + Format(Hex(i), "00"), Chr(i))
        Next i
        RemoveLocalPaths = Replace(cCode, hfPath, "")
    Else
        RemoveLocalPaths = cCode
    End If
    
    Exit Function
    
ReturnDefault:
    RemoveLocalPaths = cCode

End Function

Private Sub SaveHotSpots()

    Dim FF As Integer
    Dim cE As IHTMLElement
    Dim cs As IHTMLScriptElement
    Dim i As Integer
    Dim cCode As String
    Dim itm As IHTMLElement
    Dim doc2 As IHTMLDocument2
    Dim BodyHTML As String
    Dim eCode(1 To 3) As String
    Dim HSImageName As String
    
    Dim oCode As String
    Dim nCode As String
    Dim tCode As String
    Dim IsSameHS As Boolean
    
    On Error GoTo AbortOperation
    
    Set doc2 = wbCtrl2.Document
    
    For Each cE In doc2.body.All
        nCode = RemoveLocalPaths(LCase(cE.innerHTML))
        If nCode <> "" Then
            For i = 1 To cHotSpot.Count - 1
                oCode = RemoveLocalPaths(LCase(HotSpots(i).innerHTML))
                
                IsSameHS = False
                If oCode <> "" And nCode <> "" Then
                    If Len(oCode) > Len(nCode) Then
                        IsSameHS = (Left(nCode, Len(nCode) - 1) = Left(oCode, Len(nCode) - 1))
                    End If
                    If Len(oCode) < Len(nCode) Then
                        IsSameHS = (Left(oCode, Len(oCode) - 1) = Left(nCode, Len(oCode) - 1))
                    End If
                    If Len(oCode) = Len(nCode) Then
                        IsSameHS = (nCode = oCode)
                    End If
                End If
                
                If IsSameHS Then
                    If cE.tagName = "A" Then
                        Set itm = cE
                    Else
                        Set itm = cE.parentElement
                    End If
                    If Not cE.Children(0) Is Nothing Then
                        If cE.Children(0).tagName = "IMG" Then
                            HSImageName = cE.Children(0).Name
                        End If
                    End If
                    With itm
                        eCode(1) = RemoveOurEventCode(.OnMouseOver & "")
                        eCode(2) = RemoveOurEventCode(.onmouseout & "")
                        eCode(3) = RemoveOurEventCode(.OnClick & "")
                        cCode = GetGroupEventCode(GetIDByName(Split(cHotSpot(i).Tag, ",")(0)), HSImageName, False)
                        .setAttribute "onmouseover", GetEventCode("onmouseover", cCode) + eCode(1)
                        .setAttribute "onmouseout", GetEventCode("onmouseout", cCode) + eCode(2)
                        .setAttribute "onclick", GetEventCode("onclick", cCode) + eCode(3)
                    End With
                    Exit For
                End If
            Next i
        End If
    Next cE

    Dim pCode As String
    Dim sCode As String
    Dim BackupFile As String
    
    BackupFile = GetFileName(Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile) + ".back"
    FileCopy Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile, BackupFile

    FF = FreeFile
    Open Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile For Input As #FF
        pCode = Input$(LOF(FF), FF)
    Close #FF
    
    'Insert the modifyed trigger code
    If InStr(LCase$(pCode), "<body") = 0 Then
        pCode = TrimCR(doc2.body.innerHTML)
    Else
        pCode = Left$(pCode, InStr(InStr(LCase$(pCode), "<body"), pCode, ">")) + _
            TrimCR(doc2.body.innerHTML) + _
            Mid$(pCode, InStr(LCase$(pCode), "</body>"))
    End If
    
    'Remove any rendered JavaScript call to the loader code
    If InStr(pCode, LoaderCodeSTART) > 0 Then
        pCode = RemoveLoaderCode(pCode)
    Else
        For Each cs In doc2.scripts
            If InStr(cs.Text, "iemenu.js") And InStr(pCode, "iemenu.js") Then
                pCode = Left$(pCode, InStrRev(LCase$(pCode), "<script", InStr(pCode, cs.Text)) - 1) + Mid$(pCode, InStr(InStr(pCode, cs.Text), LCase$(pCode), "</script>") + 9)
            ElseIf InStr(cs.src, "menu.js") And InStr(pCode, "menu.js") Then
                pCode = Left$(pCode, InStrRev(LCase$(pCode), "<script", InStr(pCode, cs.src)) - 1) + Mid$(pCode, InStr(InStr(pCode, cs.src), LCase$(pCode), "</script>") + 9)
            End If
        Next cs
    End If
    
    'Get the new loader code
    sCode = TrimCR(GenLoaderCode(False, False))
    'Attach the new loader code
    pCode = AttachLoaderCode(pCode, sCode)
    
    Do While InStr(pCode, "<TBODY>") Or InStr(pCode, "</TBODY>")
        pCode = Replace(pCode, "<TBODY>", "")
        pCode = Replace(pCode, "</TBODY>", "")
    Loop

    'Save the new document
    SaveNewCode pCode, Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile
    
    Kill BackupFile
    
    GoTo CleanSave
    
AbortOperation:
    MsgBox GetLocalizedStr(618) + vbCrLf + GetLocalizedStr(544) + " " & Err.Number & ": " + Err.Description, vbCritical + vbOKOnly, GetLocalizedStr(335)

    On Error Resume Next
    If BackupFile <> "" Then
        Kill Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile
        FileCopy BackupFile, Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile
    End If
    
CleanSave:
    Set cE = Nothing
    Set cs = Nothing
    Set itm = Nothing
    Set doc2 = Nothing
    
    tmrSaveClose.Enabled = True
    
End Sub

Private Sub SaveNewCode(nCode As String, FileName As String)

    Dim FF As Integer
    Dim oCode As String
    
    Dim os() As String
    Dim ol As String
    Dim ofn As String
    Dim ns() As String
    Dim nl As String
    Dim nfn As String
    Dim i As Integer
    Dim J As Integer
    
    FF = FreeFile
    Open FileName For Input As #FF
        oCode = LCase(Input$(LOF(FF), FF))
    Close #FF
    
    os = Split(LCase(oCode), " src")
    ns = Split(LCase(nCode), " src")
    
    For i = 1 To UBound(os)
        ol = GetSRCLink(os(i))
        ofn = GetFileName(ol)
        For J = 1 To UBound(ns)
            nl = GetSRCLink(ns(J))
            nfn = GetFileName(nl)
            If (ofn = nfn) And (nfn <> "") And (ofn <> "") Then
                If nl <> ol Then
                    nCode = Replace(nCode, nl, ol)
                End If
                Exit For
            End If
        Next J
    Next i
    
    oCode = ""
    Do While nCode <> oCode
        oCode = nCode
        nCode = Replace(nCode, LoaderCodeEND + vbCrLf + vbCrLf, LoaderCodeEND + vbCrLf)
    Loop

    FF = FreeFile
    Open FileName For Output As #FF
        Print #FF, nCode
    Close #FF

End Sub

Private Function GetSRCLink(ByVal lnk As String) As String

    Dim p1 As Integer
    Dim p2 As Integer
    Dim c As String
    
    c = """"
    p1 = InStr(lnk, c)
    If InStr(lnk, "'") < p1 And InStr(lnk, "'") <> 0 Or p1 = 0 Then
        c = "'"
        p1 = InStr(lnk, c)
    End If
    
    p2 = InStr(p1 + 1, lnk, c)
    If p2 = 0 Then Err.Raise vbError + 500, "GetSRCLink", "Missing double quotes for parameter 'src'"
    
    GetSRCLink = Mid(lnk, p1 + 1, p2 - p1 - 1)

End Function

Private Function RemoveOurEventCode(ByVal pCode As String) As String

    Dim cCode As String
    
    If InStr(pCode, "function anonymous()") Then
        pCode = Mid$(pCode, InStr(pCode, "{") + 1)
        pCode = Mid$(pCode, 1, InStrRev(pCode, "}") - 1)
    End If

    'Compatibility for older versions of DHTML Menu Builder
    cCode = "if(IE){event.srcElement.style.cursor='"
    If InStr(pCode, cCode) Then
        If InStr(InStr(pCode, cCode), pCode, "';}") Then
            pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid$(pCode, InStr(InStr(pCode, cCode), pCode, "';}") + 3)
        Else
            pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid$(pCode, InStr(InStr(pCode, cCode), pCode, "'}") + 3)
        End If
    End If
    
    cCode = "cFrame.tHideAll();"
    If InStr(pCode, cCode) Then
        pCode = Trim$(TrimCR(Join(Split(pCode, cCode))))
    End If
    
    cCode = "tHideAll();"
    If InStr(pCode, cCode) Then
        pCode = Trim$(TrimCR(Join(Split(pCode, cCode))))
    End If
    
    If InStr(pCode, "cFrame.") Then
        cCode = "cFrame.ShowMenu("
        If InStr(pCode, "ShowMenu(") Then
            pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid$(pCode, InStr(InStr(pCode, cCode), pCode, ");") + 2)
        End If
    Else
        'Compatibility for older versions of DHTML Menu Builder
        cCode = "ShowMenu("
        If InStr(pCode, "ShowMenu(") Then
            pCode = Left$(pCode, InStr(pCode, cCode) - 1) + Mid$(pCode, InStr(InStr(pCode, cCode), pCode, "false)") + 6)
        End If
    End If
    
    pCode = Replace(pCode, "{;}", ";")
    pCode = Replace(pCode, "{}", "")
    While InStr(pCode, ";;") <> 0
        pCode = Replace(pCode, ";;", ";")
    Wend
    pCode = TrimCR(pCode)
    If pCode = ";" Then pCode = ""
    RemoveOurEventCode = pCode
    
End Function

Private Function TrimCR(ByVal str As String) As String

    Do While (Left$(str, 1) = vbCr) Or (Left$(str, 1) = vbCrLf) Or (Left$(str, 1) = vbLf)
        str = Mid$(str, 2)
    Loop
    
    Do While (Right$(str, 1) = vbCr) Or (Right$(str, 1) = vbCrLf) Or (Right$(str, 1) = vbLf)
        str = Left$(str, Len(str) - 1)
    Loop
    
    TrimCR = str

End Function

Private Function GetEventCode(EventName As String, ByVal dCode As String) As Variant

    Dim EventPos As Integer
    
    If InStr(LCase$(dCode), EventName) Then
        EventPos = InStr(LCase$(dCode), EventName) + Len(EventName) + 2
        dCode = Mid$(dCode, EventPos)
        GetEventCode = Left$(dCode, InStr(dCode, Chr(34)) - 1)
    Else
        GetEventCode = Null
    End If

End Function

Private Sub BuildAlignmentCombo()

    Dim nItem As ComboItem
    
    icmbAlignment.ImageList = frmMain.ilAlignment
    
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(116), 8)
    nItem.Tag = 0
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(117), 2)
    nItem.Tag = 1
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(118), 7)
    nItem.Tag = 2
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(119), 1)
    nItem.Tag = 3
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(120), 4)
    nItem.Tag = 4
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(121), 3)
    nItem.Tag = 5
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(122), 6)
    nItem.Tag = 6
    Set nItem = icmbAlignment.ComboItems.Add(, , GetLocalizedStr(123), 5)
    nItem.Tag = 7
    
    icmbAlignment.ComboItems(1).Selected = True

End Sub

Private Sub DeleteJSFiles()

    On Error Resume Next
    With Project.UserConfigs(Project.DefaultConfig)
        FileCopy .CompiledPath + "\ie" + Project.JSFileName + ".js", .CompiledPath + "\iemenu.js.back"
        FileCopy .CompiledPath + "\" + Project.JSFileName + ".js", .CompiledPath + "\menu.js.back"
        Kill .CompiledPath + "\ie" + Project.JSFileName + ".js"
        Kill .CompiledPath + "\" + Project.JSFileName + ".js"
    End With

End Sub

Private Sub RestoreJSFiles()

    On Error Resume Next
    With Project.UserConfigs(Project.DefaultConfig)
        FileCopy .CompiledPath + "\iemenu.js.back", .CompiledPath + "\ie" + Project.JSFileName + ".js"
        FileCopy .CompiledPath + "\menu.js.back", .CompiledPath + "\" + Project.JSFileName + ".js"
        Kill .CompiledPath + "\menu.js.back"
        Kill .CompiledPath + "\iemenu.js.back"
    End With

End Sub

Private Sub Form_Load()

    Dim g As Integer
    
    LocalizeUI
    
    Caption = GetLocalizedStr(335) + " - " + Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.IndexFile
    defCaption = Caption
    
    If Val(GetSetting(App.EXEName, "HSEWinPos", "X")) = 0 Then
        CenterForm Me
    Else
        Top = GetSetting(App.EXEName, "HSEWinPos", "X")
        Left = GetSetting(App.EXEName, "HSEWinPos", "Y")
        Width = GetSetting(App.EXEName, "HSEWinPos", "W")
        Height = GetSetting(App.EXEName, "HSEWinPos", "H")
        WindowState = Val(GetSetting(App.EXEName, "HSEWinPos", "State"))
    End If
    
    SetupCharset Me
    
    SetCtrlIcon
    BuildAlignmentCombo
    
    DeleteJSFiles
    
    With wbCtrl2
        .Visible = True
        .Silent = True
        .Navigate Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile
    End With
    
    For g = 1 To UBound(MenuGrps)
        cmbMenuGroups.ComboItems.Add , , MenuGrps(g).Name, IconIndex("Group")
    Next g
    If cmbMenuGroups.ComboItems.Count Then cmbMenuGroups.ComboItems(1).Selected = True
    cmbMenuGroups_Click
    
    With wbCtrl
        .Width = 1
        .Height = 1
        .Silent = True
        If Project.UserConfigs(Project.DefaultConfig).Frames.UseFrames And (Project.UserConfigs(Project.DefaultConfig).Frames.CodeFrame <> Project.UserConfigs(Project.DefaultConfig).Frames.MenuFrame) Then
            .Navigate Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile
        Else
            .Navigate Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.IndexFile
        End If
    End With
    
End Sub

Private Sub SetCtrlIcon()

    cmbMenuGroups.ImageList = frmMain.ilIcons
    tbOptions.ImageList = frmMain.ilIcons
    tbOptions.Buttons("tbNext").Image = IconIndex("Redo")
    tbOptions.Buttons("tbHelp").Image = IconIndex("Help")

End Sub
    
Private Sub HighlightNewHotSpots()

    Dim i As Integer
    Dim x, Y As Long
    
    If IsBusy Then Exit Sub
    IsBusy = True
    
    For i = 1 To cHotSpot.Count - 1
        With cHotSpot(i)
            .Left = Split(.Tag, ",")(1) - ofLeft * 15
            .Top = Split(.Tag, ",")(2) - ofTop * 15
            .Width = Split(.Tag, ",")(3)
            .Height = Split(.Tag, ",")(4)
            .Cls
            .Visible = False
            Refresh
            DoEvents
            BitBlt cHotSpot(i).hDC, 0, 0, .Width / Screen.TwipsPerPixelX, .Height / Screen.TwipsPerPixelY, ChDC, .Left / Screen.TwipsPerPixelX, .Top / Screen.TwipsPerPixelY, vbSrcCopy
            BlurHotSpot i, Split(.Tag, ",")(0) = cmbMenuGroups.Text
            .ZOrder 0
            .Visible = True
        End With
    Next i
    
ExitSub:
    IsBusy = False

End Sub

Private Sub HighlightHotSpots()

    Dim cE As IHTMLElement
    Dim cE2 As IHTMLElement2
    Dim bRect As IHTMLRect
    Dim tEvent As String

    For Each cE In doc.All.tags("A")
        tEvent = ""
        If cE.getAttribute("onmouseover") <> "" Then
            If InStr(cE.getAttribute("onmouseover"), "ShowMenu") Then
                tEvent = "onmouseover"
            ElseIf InStr(cE.getAttribute("onclick"), "ShowMenu") Then
                tEvent = "onclick"
            End If
        End If
        
        If tEvent <> "" Then
            FoundAttachedMenus = FoundAttachedMenus + 1
            Set cE2 = cE
            Set bRect = cE2.getBoundingClientRect
            
            Load cHotSpot(cHotSpot.Count)
            ReDim Preserve HotSpots(cHotSpot.Count)
            
            With cHotSpot(cHotSpot.Count - 1)
                .Left = bRect.Left * 15 + wbCtrl.Left
                .Top = bRect.Top * 15 + wbCtrl.Top
                .Width = (bRect.Right - bRect.Left) * 15
                .Height = (bRect.Bottom - bRect.Top) * 15
                BitBlt cHotSpot(cHotSpot.Count - 1).hDC, 0, 0, .Width / Screen.TwipsPerPixelX, .Height / Screen.TwipsPerPixelY, ChDC, .Left / Screen.TwipsPerPixelX, .Top / Screen.TwipsPerPixelY, vbSrcCopy
                .Tag = Mid$(cE.getAttribute(tEvent), InStr(cE.getAttribute(tEvent), "ShowMenu") + 10)
                .Tag = Left$(.Tag, InStr(.Tag, "'") - 1) & "," & .Left & "," & .Top & "," & .Width & "," & .Height & "," & cE.outerHTML
                .ToolTipText = "Attached Group: " + Split(.Tag, ",")(0)
                If GetIDByName(Split(.Tag, ",")(0)) = 0 Then
                    'A group has been deleted or renamed and the corresponding HotSpot must be removed
                    ReDim Preserve HotSpots(UBound(HotSpots) - 1)
                    Unload cHotSpot(cHotSpot.Count - 1)
                    FoundAttachedMenus = FoundAttachedMenus - 1
                Else
                    Set HotSpots(cHotSpot.Count - 1) = cE
                    BlurHotSpot cHotSpot.Count - 1, Split(.Tag, ",")(0) = cmbMenuGroups.Text
                    .ZOrder 0
                    .Visible = True
                End If
            End With
        End If
    Next cE
    
    Set cE = Nothing
    Set cE2 = Nothing
    Set bRect = Nothing
    
End Sub

Private Sub BlurHotSpot(hIdx As Integer, Highlighted As Boolean)

    Dim x As Long
    Dim Y As Long
    Dim n As Integer
    Dim cp As Long

    For Y = 1 To cHotSpot(hIdx).Height Step 15
        n = IIf(Int(Y / 2) = Y / 2, 0, 1)
        For x = 1 To cHotSpot(hIdx).Width Step 15
            n = n + 1
            cp = cHotSpot(hIdx).Point(x, Y) And &HFF
            If cp = 0 Then cp = 128
            If Int(n / 2) = n / 2 Then
                If Highlighted Then
                    cHotSpot(hIdx).PSet (x, Y), RGB(cp / 2, cp / 2, cp)
                Else
                    cHotSpot(hIdx).PSet (x, Y), RGB(cp / 1.5, cp / 1.5, cp / 1.5)
                End If
            End If
    Next x, Y

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    If WindowState = vbNormal Then
        SaveSetting App.EXEName, "HSEWinPos", "X", Top
        SaveSetting App.EXEName, "HSEWinPos", "Y", Left
        SaveSetting App.EXEName, "HSEWinPos", "W", Width
        SaveSetting App.EXEName, "HSEWinPos", "H", Height
    End If
    SaveSetting App.EXEName, "HSEWinPos", "State", WindowState

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    cmdClose.Move Width - 1170, Height - 810
    cmdSave.Move Width - 2280, Height - 810

    wbCtrl.Move 45, tbOptions.Top + tbOptions.Height + 30, Width - 235, Height - cmdSave.Height - tbOptions.Top - tbOptions.Height - 600
    
    DoEvents
    
    HighlightNewHotSpots

End Sub

Private Sub CheckHotSpotSetup(cE As IHTMLElement, g As Integer)

    On Error GoTo ExitSub

    If cE.Children(0).tagName = "IMG" Then
        If cE.Children(0).Name = "" Then
            MsgBox GetLocalizedStr(619) + vbCrLf + _
                   GetLocalizedStr(620) + " " + cmbMenuGroups.Text + " " + GetLocalizedStr(621) + vbCrLf + _
                   GetLocalizedStr(622), vbInformation + vbOKOnly, GetLocalizedStr(623)
            Exit Sub
        End If
        If MenuGrps(g).Name = cE.Children(0).Name Then
            MsgBox GetLocalizedStr(624) + vbCrLf + _
                   GetLocalizedStr(625), vbInformation + vbOKOnly, GetLocalizedStr(623)
            Exit Sub
        End If
        MenuGrps(g).HSImage = cE.Children(0).Name
    End If
    
ExitSub:

End Sub

Private Sub HighlightCurrent()

    Dim i As Integer
    Dim x, Y As Long
    Dim hX, hY As Long
    Dim g As Integer
    
    Project.HasChanged = True
    ChangesSaved = False
    
    For i = 1 To cHotSpot.Count - 1
        If Split(cHotSpot(i).Tag, ",")(0) = cmbMenuGroups.Text Then
            Set HotSpots(i) = cEInfo
            With cHotSpot(i)
                .Cls
                .Visible = False
                Refresh
                DoEvents
                .Left = picElement.Left
                .Top = picElement.Top
                .Width = picElement.Width
                .Height = picElement.Height
                BitBlt cHotSpot(i).hDC, 0, 0, .Width / Screen.TwipsPerPixelX, .Height / Screen.TwipsPerPixelY, ChDC, .Left / Screen.TwipsPerPixelX, .Top / Screen.TwipsPerPixelY, vbSrcCopy
                .Tag = Split(.Tag, ",")(0) & "," & .Left + ofLeft * 15 & "," & .Top + ofTop * 15 & "," & .Width & "," & .Height
                BlurHotSpot i, True
                
                g = GetIDByName(Split(.Tag, ",")(0))
                CheckHotSpotSetup cEInfo, g
                
                With MenuGrps(g)
                    .x = Fix((picMenuShape.Left - wbCtrl.Left + ofLeft) / Screen.TwipsPerPixelX) + ofLeft - 1
                    .Y = Fix((picMenuShape.Top - wbCtrl.Top) / Screen.TwipsPerPixelY) + ofTop - 2
                    If Project.UserConfigs(Project.DefaultConfig).Frames.UseFrames And (Project.UserConfigs(Project.DefaultConfig).Frames.CodeFrame <> Project.UserConfigs(Project.DefaultConfig).Frames.MenuFrame) Then
                        Select Case Val(icmbAlignment.SelectedItem.Tag)
                            Case 0  ' Bottom Left
                                .Y = 0
                            Case 1  ' Bottom Right
                                .Y = 0
                            Case 2  ' Top Left
                            Case 3  ' Top Right
                            Case 4  ' Left Top
                                .x = 0
                            Case 5  ' Right Top
                        End Select
                    End If
                End With
                
                .ZOrder 0
                .Visible = True
            End With
            Exit Sub
        End If
    Next i
    
    Load cHotSpot(cHotSpot.Count)
    On Error GoTo CreateHotSpotObject
    Set HotSpots(cHotSpot.Count - 1) = cEInfo
    
    With cHotSpot(cHotSpot.Count - 1)
        .Cls
        .Visible = False
        Refresh
        DoEvents
        .Left = picElement.Left
        .Top = picElement.Top
        .Width = picElement.Width
        .Height = picElement.Height
        BitBlt cHotSpot(cHotSpot.Count - 1).hDC, 0, 0, .Width / Screen.TwipsPerPixelX, .Height / Screen.TwipsPerPixelY, ChDC, .Left / Screen.TwipsPerPixelX, .Top / Screen.TwipsPerPixelY, vbSrcCopy
        .Tag = cmbMenuGroups.Text & "," & .Left + ofLeft * 15 & "," & .Top + ofTop * 15 & "," & .Width & "," & .Height
        .ToolTipText = "Attached Group: " + Split(.Tag, ",")(0)
        BlurHotSpot cHotSpot.Count - 1, True
        
        g = GetIDByName(Split(.Tag, ",")(0))
        CheckHotSpotSetup cEInfo, g
        
        With MenuGrps(g)
            .x = Fix((picMenuShape.Left - wbCtrl.Left + ofLeft) / Screen.TwipsPerPixelX) + ofLeft - 1
            .Y = Fix((picMenuShape.Top - wbCtrl.Top) / Screen.TwipsPerPixelY) + ofTop - 2
        End With
                
        .ZOrder 0
        .Visible = True
    End With
    
    Exit Sub
    
CreateHotSpotObject:
    ReDim Preserve HotSpots(cHotSpot.Count - 1)
    Resume
    
End Sub

Private Sub icmbAlignment_Click()

    MenuGrps(cmbMenuGroups.SelectedItem.Index).Alignment = Val(icmbAlignment.SelectedItem.Tag)

End Sub

Private Sub picElement_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)

    If Button = vbLeftButton Then
        picElement.Visible = False
        Refresh
        DoEvents
        HighlightCurrent
        picElement.Visible = True
    End If

End Sub

Private Sub tbOptions_ButtonClick(ByVal Button As MSComctlLib.Button)

    Select Case Button.key
        Case "tbNext"
            If cmbMenuGroups.SelectedItem.Index = cmbMenuGroups.ComboItems.Count Then
                cmbMenuGroups.ComboItems.item(1).Selected = True
            Else
                cmbMenuGroups.ComboItems.item(cmbMenuGroups.SelectedItem.Index + 1).Selected = True
            End If
            cmbMenuGroups_Click
        Case "tbHelp"
            ShowHelp "hotspot_editor.htm"
    End Select

End Sub

Private Sub tmrDoneLoading_Timer()

    tmrDoneLoading.Enabled = False

    HighlightHotSpots
    Form_Resize
    tmrMonitor.Enabled = True
    tmrRefresh.Enabled = True

End Sub

Private Sub tmrMonitor_Timer()

    Dim crsrPos As POINTAPI
    Dim cElement As IHTMLElement2
    Dim cEScroll As IHTMLElement2
    Dim ceChildren As IHTMLElement
    Dim areaChildren As IHTMLDOMChildrenCollection
    Dim cERect As IHTMLRect
    Dim l, t, w, h As Long
    Static ll, lt, lw, lh As Long
    
    If MonitorBusy Then Exit Sub
    MonitorBusy = True
    
    If Not doc Is Nothing Then
        Set cEScroll = doc.body
        With cEScroll
            If ofLeft <> .scrollLeft Or ofTop <> .scrollTop Then
                ofLeft = .scrollLeft
                ofTop = .scrollTop
                HighlightNewHotSpots
            End If
        End With
    End If
    
    GetCursorPos crsrPos
    my = crsrPos.Y - (Top + wbCtrl.Top) / Screen.TwipsPerPixelX - 24
    mx = crsrPos.x - (Left + wbCtrl.Left) / Screen.TwipsPerPixelY - 3
    
    If doc.elementFromPoint(0, 0) Is Nothing Then
        If Not ConLost Then
            ConLost = True
            MsgBox GetLocalizedStr(626), vbCritical + vbOKOnly, GetLocalizedStr(335)
        End If
    End If
    
    Set cElement = doc.elementFromPoint(mx, my)
    
    If Not cElement Is Nothing Then
        Set cEInfo = FindAnchorTag(cElement)
        If Not cEInfo Is Nothing Then
            If cEInfo.tagName = "AREA" Then
                Set cElement = GetMappedImage(cEInfo.parentElement.getAttribute("name"))
            End If
            Set cERect = cElement.getBoundingClientRect
            With cERect
                l = .Left + wbCtrl.Left / Screen.TwipsPerPixelX
                t = .Top + wbCtrl.Top / Screen.TwipsPerPixelY
                w = .Right - .Left
                h = .Bottom - .Top
            End With
            
            If Not picMenuShape.Visible Then
                ll = l: lt = t: lw = w: lh = h
                If w + l >= Fix(wbCtrl.Width / Screen.TwipsPerPixelX) Then w = Fix((wbCtrl.Width - wbCtrl.Left) / Screen.TwipsPerPixelX) - 12
                If t <= Fix(wbCtrl.Top / Screen.TwipsPerPixelY) Then t = Fix(wbCtrl.Top / Screen.TwipsPerPixelY)
                If h + t >= Fix(wbCtrl.Height / Screen.TwipsPerPixelY) Then h = Fix(wbCtrl.Height / Screen.TwipsPerPixelY) - t + 22
                
                picElement.Visible = False
                picMenuShape.Visible = False
                Refresh
                DoEvents
                On Error Resume Next
                picElement.Move l * 15, t * 15, w * 15, h * 15
                picElement.Cls
                BitBlt picElement.hDC, 0, 0, w, h, ChDC, l, t, vbSrcInvert
            End If
            picElement.Visible = True
            picElement.ZOrder 0
            picMenuShape.Visible = True
        Else
            picElement.Visible = False
            picMenuShape.Visible = False
        End If
        
        With picMenuShape
            If Not icmbAlignment.SelectedItem Is Nothing Then
                Select Case Val(icmbAlignment.SelectedItem.Tag)
                    Case gacBottomLeft
                        .Left = picElement.Left
                        .Top = picElement.Top + picElement.Height
                    Case gacBottomRight
                        .Left = picElement.Left + picElement.Width - .Width
                        .Top = picElement.Top + picElement.Height
                    Case gacTopLeft
                        .Left = picElement.Left
                        .Top = picElement.Top - .Height
                    Case gacTopRight
                        .Left = picElement.Left + picElement.Width - .Width
                        .Top = picElement.Top - .Height
                    Case gacLeftTop
                        .Left = picElement.Left - .Width
                        .Top = picElement.Top
                    Case gacLeftBottom
                        .Left = picElement.Left - .Width
                        .Top = picElement.Top + picElement.Height - .Height
                    Case gacRightTop
                        .Left = picElement.Left + picElement.Width
                        .Top = picElement.Top
                    Case gacRightBottom
                        .Left = picElement.Left + picElement.Width
                        .Top = picElement.Top + picElement.Height - .Height
            End Select
            End If
            On Error Resume Next
'            If (.Left + .Width) > (wbCtrl.Left + wbCtrl.Width) Then .Width = wbCtrl.Width - .Left
'            If .Left < wbCtrl.Left Then .Left = wbCtrl.Left: .Width = .Width - (wbCtrl.Left - .Left)
'            If (.Top + .Height) > (wbCtrl.Top + wbCtrl.Height) Then .Height = wbCtrl.Height - .Top
'            If .Top < wbCtrl.Top Then .Top = wbCtrl.Top: .Height = .Height - (wbCtrl.Top - .Top)
            SetRectShape picMenuShape, cmbMenuGroups.SelectedItem.Index
            .ZOrder 0
        End With
    End If
    
    Set cElement = Nothing
    Set cEScroll = Nothing
    Set ceChildren = Nothing
    Set areaChildren = Nothing
    Set cERect = Nothing
    
    MonitorBusy = False
    
End Sub

Private Function FindAnchorTag(cE As IHTMLElement) As IHTMLElement

    Do Until cE Is Nothing
        If cE.tagName = "A" Then
            Exit Do
        End If
        If cE.tagName = "HTML" Then
            Set cE = Nothing
            Exit Do
        End If
        Set cE = cE.parentElement
    Loop
    
    Set FindAnchorTag = cE

End Function

Private Function GetMappedImage(MapName As String) As IHTMLElement2

    Dim tmpEl As IHTMLElement
    
    For Each tmpEl In doc.All
        If tmpEl.tagName = "IMG" Then
            If tmpEl.getAttribute("usemap") <> "" Then
                Set GetMappedImage = tmpEl
                Exit For
            End If
        End If
    Next tmpEl
    
    Set tmpEl = Nothing

End Function

Private Sub tmrMonitorLoad_Timer()

    Dim t As Integer

    t = CInt((pg1 + pg2) / 2)
    If t < 100 Then
        Caption = defCaption + " (" & t & "%)"
    Else
        Caption = defCaption
        tmrMonitorLoad.Enabled = False
    End If

End Sub

Private Sub tmrRefresh_Timer()

    If Not picMenuShape.Visible Then HighlightNewHotSpots
    'tmrRefresh.Interval = Int((5000 - 1000 + 1) * Rnd + 1000)
    tmrRefresh.Enabled = False

End Sub

Private Sub tmrSaveClose_Timer()

    Dim i As Integer
    
    On Error Resume Next

    tmrSaveClose.Enabled = False
    If Not ChangesSaved Then
        If MsgBox(GetLocalizedStr(627), vbQuestion + vbYesNo, GetLocalizedStr(628)) = vbNo Then
            tmrMonitor.Enabled = True
            tmrRefresh.Enabled = True
            
            DoEvents
        
            cmdSave.Enabled = True
            cmdClose.Enabled = True
            MousePointer = vbDefault
            Exit Sub
        End If
    End If
    
    Me.Enabled = False
    
    RestoreJSFiles

    ReleaseDC hWnd, ChDC
    Set doc = Nothing
    Set cEInfo = Nothing
    
    'For i = 0 To UBound(HotSpots) - 1
    '    Set HotSpots(i) = Nothing
    'Next i
    
    Caption = "Closing"
    While (wbCtrl.Busy Or wbCtrl2.Busy) Or (wbCtrl.ReadyState = READYSTATE_LOADING Or wbCtrl2.ReadyState = READYSTATE_LOADING)
        Caption = Caption + "."
        DoEvents
        wbCtrl.Document.body.innerHTML = ""
        wbCtrl2.Document.body.innerHTML = ""
        DoEvents
    Wend
    
    Unload Me

End Sub

Private Sub tmrWait_Timer()

    Static intCount As Integer
    
    intCount = intCount + 1

    If (MonitorBusy Or IsBusy) And intCount < 200 Then
        DoEvents
        Exit Sub
    End If
    
    tmrWait.Enabled = False
    
    Select Case UserAction
        Case uaSAVE
            InitSave
        Case uaCLOSE
            InitClose
    End Select

End Sub

Private Sub wbCtrl_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, Flags As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

    Cancel = DisableNav

End Sub

Private Sub wbCtrl_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    Dim pWin As IHTMLWindow2
    Dim g As Integer
    Dim c As Integer
    Dim IsSubGrp As Boolean
    Dim numGrps As Integer
    Dim i As Integer
    
    If LCase(URL) <> LCase(Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.IndexFile) Then Exit Sub
    
    Do While wbCtrl.Document Is Nothing
        DoEvents
        i = i + 1
        If i > 32700 Then Exit Sub
    Loop

    Set doc = wbCtrl.Document
    Set pWin = doc.parentWindow
    
    pWin.execScript "IE=NS=false;function ShowMenu() {;}"
    
    ChDC = GetDC(hWnd)
    
    tmrDoneLoading.Enabled = True
    
'    If AutoHotSpot Then
'        AutoHotSpot = False
'        For g = 1 To UBound(MenuGrps)
'            IsSubGrp = False
'            For c = 1 To UBound(MenuCmds)
'                If MenuCmds(c).Actions.OnClick.Type = atcCascade And MenuCmds(c).Actions.OnClick.TargetMenu = g Or _
'                    MenuCmds(c).Actions.OnMouseOver.Type = atcCascade And MenuCmds(c).Actions.OnMouseOver.TargetMenu = g Or _
'                    MenuCmds(c).Actions.OnDoubleClick.Type = atcCascade And MenuCmds(c).Actions.OnDoubleClick.TargetMenu = g Then
'                    IsSubGrp = True
'                    Exit For
'                End If
'            Next c
'            If Not IsSubGrp Then numGrps = numGrps + 1
'        Next g
'        If FoundAttachedMenus < numGrps Then
'            If MsgBox(GetLocalizedStr(629), vbQuestion + vbYesNo, GetLocalizedStr(630)) = vbNo Then
'                cmdClose_Click
'            End If
'        Else
'            cmdSave_Click
'        End If
'    End If
    
End Sub

Private Sub SetRectShape(pic As PictureBox, GroupId As Integer)

    pic.Width = GetDivWidth(GroupId) * 15
    pic.Height = GetDivHeight(GroupId) * 15
    pic.Tag = (GetDivWidth(GroupId) * 15) & "|" & (GetDivHeight(GroupId) * 15)
    pic.BackColor = MenuGrps(GroupId).bColor
    pic.BorderStyle = Abs((MenuGrps(GroupId).FrameBorder > 0))

End Sub

Private Sub wbCtrl_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

    On Error Resume Next
    
    pg1 = (Progress / ProgressMax) * 100

End Sub

Private Sub wbCtrl2_DocumentComplete(ByVal pDisp As Object, URL As Variant)

    Dim pWin2 As IHTMLWindow2
    Dim doc2 As IHTMLDocument2
    Dim i As Integer

    If LCase(URL) <> LCase(Project.UserConfigs(Project.DefaultConfig).HotSpotEditor.HotSpotsFile) Then Exit Sub

    i = 0
    Do While wbCtrl2.Document Is Nothing
        DoEvents
        i = i + 1
        If i > 32700 Then Exit Sub
    Loop

    Set doc2 = wbCtrl2.Document
    Set pWin2 = doc2.parentWindow

    pWin2.execScript "IE=NS=false;function ShowMenu() {;}"

    wbCtrl2.Visible = False

End Sub

Private Sub LocalizeUI()

    cmdSave.Caption = GetLocalizedStr(130)
    cmdClose.Caption = GetLocalizedStr(424)
    
    lblGroups.Caption = GetLocalizedStr(505)
    lblAlignment.Caption = GetLocalizedStr(115)
    
    tbOptions.Buttons("tbNext").Caption = GetLocalizedStr(616)
    
    If Preferences.Language <> "eng" Then
        cmdSave.Width = SetCtrlWidth(cmdSave)
        cmdClose.Width = SetCtrlWidth(cmdClose)
    End If

End Sub

Private Sub wbCtrl2_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)

    On Error Resume Next
    
    pg2 = (Progress / ProgressMax) * 100
    
End Sub

#End If
