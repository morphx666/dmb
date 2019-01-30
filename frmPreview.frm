VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{E1C6DB9D-BD4A-4A61-A759-0CED75D034BF}#43.0#0"; "SmartButton.ocx"
Begin VB.Form frmPreview 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Preview"
   ClientHeight    =   6225
   ClientLeft      =   3135
   ClientTop       =   2685
   ClientWidth     =   7560
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPreview.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   6225
   ScaleWidth      =   7560
   Begin VB.Frame frameControls 
      Height          =   1050
      Left            =   285
      OLEDropMode     =   1  'Manual
      TabIndex        =   1
      Top             =   4005
      Width           =   5085
      Begin VB.ComboBox cmbCharset 
         Height          =   315
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   225
         Width           =   2565
      End
      Begin VB.CheckBox chkUseTarget 
         Caption         =   "Use Target Document"
         Height          =   210
         Left            =   105
         TabIndex        =   6
         Top             =   675
         Width           =   2220
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "Refresh"
         Height          =   345
         Left            =   3210
         TabIndex        =   7
         Top             =   585
         Width           =   780
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   345
         Left            =   4185
         TabIndex        =   8
         Top             =   585
         Width           =   780
      End
      Begin SmartButtonProject.SmartButton cmdColor 
         Height          =   240
         Left            =   1440
         TabIndex        =   3
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
      Begin VB.Label lblEncoding 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Encoding"
         Height          =   195
         Left            =   2100
         TabIndex        =   4
         Top             =   285
         Width           =   645
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   1935
         X2              =   1935
         Y1              =   225
         Y2              =   540
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         X1              =   1920
         X2              =   1920
         Y1              =   225
         Y2              =   540
      End
      Begin VB.Label lblBackColor 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Background Color"
         Height          =   195
         Left            =   105
         TabIndex        =   2
         Top             =   285
         Width           =   1260
      End
   End
   Begin VB.Timer tmrForceLoad 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   540
      Top             =   5235
   End
   Begin SHDocVwCtl.WebBrowser wBrowser 
      Height          =   3195
      Left            =   60
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   420
      Width           =   4320
      ExtentX         =   7620
      ExtentY         =   5636
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
   Begin VB.Timer tmrLoadPreview 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   1350
      Top             =   5325
   End
   Begin MSComctlLib.StatusBar sbSizer 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   9
      Top             =   5970
      Width           =   7560
      _ExtentX        =   13335
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
      EndProperty
   End
End
Attribute VB_Name = "frmPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim PreviewFile As String
Dim IsLoading As Boolean

Private Function GetScriptCode() As String

    GetScriptCode = "<script language=""JavaScript"" type=""text/javascript"">" + _
                    LoadFile(PreviewPath + "menu.js") + _
                    "</script>"

End Function

Private Function TargetFileExists() As Boolean

    If FileExists(GetRealLocal.HotSpotEditor.HotSpotsFile) Then
        TargetFileExists = True
    Else
        MsgBox "The target document '" + GetRealLocal.HotSpotEditor.HotSpotsFile + "' specified on the default configuration does not exist", vbInformation + vbOKOnly, "Unable to load Target Document"
        TargetFileExists = False
    End If

End Function

Private Sub LoadPreview()

    Dim ff_template As Integer
    Dim ff_preview As Integer
    Dim sStr As String
    Dim y As Single
    Dim sCode As String
    Dim g As Integer
    Dim t As Integer
    Dim UseAttachTo As Integer
    Dim oProject As ProjectDef
    Dim hsOffset As Integer
    
    On Error GoTo chkError
    
    caption = Project.Name + " - " + GetLocalizedStr(959)
    
    ConfigIE
    chkUseTarget.Enabled = FileExists(GetRealLocal.HotSpotEditor.HotSpotsFile)
    
    If Preferences.ShowCleanPreview Then
        hsOffset = -20
    Else
        hsOffset = 20
    End If
    
    If PreviewMode = pmcNormal Then
        PreviewFile = PreviewPath + "index.html"
        
        sCode = GetScriptCode
        y = 1
        
        DoEvents
        
        For t = 1 To UBound(Project.Toolbars)
            If LenB(Project.Toolbars(t).AttachTo) <> 0 Then
                UseAttachTo = t
                Exit For
            End If
        Next t
        If UseAttachTo > 0 Then
            With TipsSys
                Do While .IsVisible
                    DoEvents
                Loop
                .CanDisable = True
                .DialogTitle = "Preview"
                .TipTitle = "'Attach To' Problems in the Preview"
                .Tip = "When using the 'Attach to' feature in one or more toolbars, the preview will not show the proper alignment on your toolbars because the reference image (or object) will not be present on the Preview; however, the toolbars using the 'Attach to' feature will be properly positioned when viewed from your web site."
                .Show
            End With
        End If
        
        If GetSetting(App.EXEName, "PreviewWinPos", "UseTarget", 0) = 1 Then
            If TargetFileExists Then
                oProject = Project
                Project.CodeOptimization = cocDEBUG
                Project.DefaultConfig = GetConfigID(GetRealLocal.Name)
                sStr = LoadFile(GetRealLocal.HotSpotEditor.HotSpotsFile)
                CompileProject MenuGrps, MenuCmds, Project, Preferences, params, True, True
                sStr = AttachLoaderCode(RemoveLoaderCode(sStr), GenLoaderCode(False, True))
                sStr = Replace(sStr, Project.JSFileName + ".js", "menu.js")
                
                sStr = FixExtRefs(sStr, GetRealLocal.RootWeb, GetFilePath(GetRealLocal.HotSpotEditor.HotSpotsFile), "img", "src")
                sStr = FixExtRefs(sStr, GetRealLocal.RootWeb, GetFilePath(GetRealLocal.HotSpotEditor.HotSpotsFile), "link", "href")
                sStr = FixExtRefs(sStr, GetRealLocal.RootWeb, GetFilePath(GetRealLocal.HotSpotEditor.HotSpotsFile), "embed", "src")
                
                SaveFile PreviewFile, sStr
                Project = oProject
            Else
                chkUseTarget.Value = vbUnchecked
                Exit Sub
            End If
        Else
            ff_preview = FreeFile
            Open PreviewFile For Output As #ff_preview
            ff_template = FreeFile
            Open AppPath + "rsc\preview.tpl" For Input As #ff_template
                Do Until EOF(ff_template)
                    Line Input #ff_template, sStr
                    sStr = Replace(sStr, "|SCRIPT|", sCode)
                    sStr = Replace(sStr, "|DMBVER|", DMBVersion)
                    sStr = Replace(sStr, "|PROJNAME|", Project.Name)
                    sStr = Replace(sStr, "|PREVIEWHELP|", AppPath + "Help\previewing.htm")
                    sStr = Replace(sStr, "|WEBCHARSET|", cs(cmbCharset.ListIndex + 1).WebCharset)
                    
                    If InStr(sStr, "|PREVIEWINFO|") Then
                        sStr = Replace(sStr, "|PREVIEWINFO|", vbNullString)
                        sStr = sStr + "<br><br><br><br><br><br><hr noshade size=1 color=#000080>"
                        If IsDEMO Then
                            sStr = sStr + "<b>The links on your menu items have been disabled on this DEMO version.</b>"
                        Else
                            sStr = sStr + "If your menu items have links, you'll get an error when you click on them. This is normal and they will work fine when used on your web page."
                        End If
                        sStr = sStr + "<br><br>"
                    End If
                    
                    If UseAttachTo > 0 Then
                        If InStr(sStr, "|IMGNAME|") > 0 Then
                            sStr = Replace(sStr, "|IMGNAME|", Project.Toolbars(UseAttachTo).AttachTo)
                        End If
                    End If
                    
                    If InStr(sStr, "|HOTSPOTS|") > 0 Then
                        sStr = Replace(sStr, "|HOTSPOTS|", vbNullString)
                        For g = 1 To UBound(MenuGrps)
                            If MemberOf(g) = 0 Or Not CreateToolbar Then
                                If Not IsSubMenu(g) Then
                                    sStr = sStr + "<tr><td width=""100%"" bgcolor=""#DCDCDC"">"
                                    sStr = sStr + "<a class=""dmbGroup"" href=""#"""
                                    t = 0
                                    If MenuGrps(g).Actions.onmouseover.Type = atcCascade Then
                                        t = MenuGrps(g).Actions.onmouseover.TargetMenu
                                    Else
                                        If MenuGrps(g).Actions.onclick.Type = atcCascade Then
                                            t = MenuGrps(g).Actions.onclick.TargetMenu
                                        End If
                                    End If
                                    If t > 0 Then
                                        sStr = sStr + " onmouseover=ShowMenu('" + MenuGrps(t).Name + "'," & 170 & "," & Int(y * 25 + hsOffset) & ",false)"
                                    End If
                                    sStr = sStr + " onmouseout=tHideAll()"
                                    sStr = sStr + ">" + MenuGrps(g).Name + "</a>"
                                    sStr = sStr + "</td></tr>"
                                    y = y + 1
                                End If
                            End If
                        Next g
                    End If
                    Print #ff_preview, sStr
                Loop
            Close #ff_template
            Close #ff_preview
        End If
        
        CopyProjectImages PreviewPath
        
        If Preferences.ShowCleanPreview Then
            If chkUseTarget.Value = vbUnchecked Or chkUseTarget.Enabled = False Then
                sStr = LoadFile(PreviewFile)
                sStr = Left(sStr, InStr(sStr, "<!-- Title START -->") - 1) + Mid(sStr, InStr(sStr, "<!-- Title END -->") + 19)
                sStr = Left(sStr, InStr(sStr, "<!-- Info START -->") - 1) + Mid(sStr, InStr(sStr, "<!-- Info END -->") + 18)
                SaveFile PreviewFile, sStr
            End If
        End If
    Else
        PreviewFile = TempPath + "smdmbp.htm"
        chkUseTarget.Enabled = False
    End If
    
    sStr = ""
    For t = 1 To UBound(Project.Toolbars)
        If Project.Toolbars(t).Alignment = tbacFree Then
            sStr = sStr + "Placeholder for toolbar #" & t & "<br><span id=dmbTB" + CStr(t) + "ph></span><p>&nbsp;</p>" + vbCrLf
        End If
    Next t
    If sStr <> "" Then
        SaveFile PreviewFile, Replace(LoadFile(PreviewFile), "</body>", sStr + "</body>", , , vbTextCompare)
    End If
    
    For t = 1 To UBound(Project.Toolbars)
        If Project.Toolbars(t).FollowVScroll Then
            sStr = ""
            For g = 1 To 32
                sStr = sStr + "<p>&nbsp;</p>" + vbCrLf
            Next g
            SaveFile PreviewFile, Replace(LoadFile(PreviewFile), "</body>", sStr + "</body>", , , vbTextCompare)
            Exit For
        End If
    Next t
    
    If Val(GetSetting(App.EXEName, "Browsers", "Default", 1)) = 1 Then
        wBrowser.Navigate Long2Short(PreviewFile)
    Else
        Dim DefBrowser As String
        DefBrowser = GetSetting(App.EXEName, "Browsers", "Command" & Val(GetSetting(App.EXEName, "Browsers", "Default", 1)) - 1)
        Shell DefBrowser + " " + Long2Short(PreviewFile), vbNormalFocus
        Unload Me
    End If
    
    Exit Sub
    
chkError:
    MsgBox "Error: " & Err.number & vbCrLf & Err.Description, vbCritical + vbOKOnly, "Error Creating Preview"
    On Error Resume Next
    Close #ff_template
    Close #ff_preview
    
End Sub

Private Function FixExtRefs(ByVal sCode As String, rw As String, ptf As String, ByVal tagName As String, ByVal paramName As String) As String

    Dim img() As String
    Dim i As Long
    Dim sp As String
        
    tagName = "<" + LCase(tagName) + " "
    
    img = Split(Replace(sCode, UCase(tagName), tagName), tagName)
    
    For i = 1 To UBound(img)
        sp = GetParamVal(img(i), paramName)
        sp = SetSlashDir(sp, sdBack)
        If InStr(sp, "dmb_i.gif") = 0 And InStr(sp, "dmb_m.gif") = 0 Then
            If Left(sp, 1) = "\" Then
                sp = Replace(rw + GetFilePath(sp) + GetFileName(sp), "\\", "\")
            Else
                sp = Replace(ptf + GetFilePath(sp) + GetFileName(sp), "\\", "\")
            End If
        img(i) = ChangeParamVal(img(i), paramName, AddFileProtocol(sp))
        End If
    Next i
    
    FixExtRefs = Join(img, tagName)

End Function

Private Sub chkUseTarget_Click()

    cmbCharset.Enabled = (chkUseTarget.Value = vbUnchecked)

    If IsLoading Then Exit Sub
    
    If chkUseTarget.Value = vbChecked Then
        If Not TargetFileExists Then
            IsLoading = True
            chkUseTarget.Value = vbUnchecked
            IsLoading = False
            Exit Sub
        End If
    End If

    SaveSetting App.EXEName, "PreviewWinPos", "UseTarget", Abs(chkUseTarget.Value = vbChecked)
    cmdRefresh_Click

End Sub

Private Sub cmbCharset_Click()

    If IsLoading Then Exit Sub
    
    Preferences.CodePage = cs(cmbCharset.ListIndex + 1).CodePage
    
    LivePreviewCharset = ""
    If Preferences.UseLivePreview Then frmMain.DoDelayedLivePreview
    
    
    cmdRefresh_Click

End Sub

Private Sub cmdClose_Click()

    Unload Me

End Sub

Private Sub cmdColor_Click()

    BuildUsedColorsArray

    With cmdColor
        SelColor = .tag
        SelColor_CanBeTransparent = False
        frmColorPicker.Show vbModal, Me
        SetColor SelColor, cmdColor
    End With
    
    SetDocBackColor
    frmMain.DoLivePreview
    
    Me.SetFocus

End Sub

Private Sub SetDocBackColor()

    On Error Resume Next
    DoEvents
    wBrowser.Document.body.bgColor = GetRGB(cmdColor.tag, True)

End Sub

Private Sub cmdRefresh_Click()

    frmMain.ShowPreview

End Sub

Private Sub Form_Load()

    IsLoading = True
    
    wBrowser.Navigate "about:blank"
    
    SetupCharset Me
    LocalizeUI
    
    PreviewIsOn = True

    If Val(GetSetting(App.EXEName, "PreviewWinPos", "X")) = 0 Then
        CenterForm Me
    Else
        Top = GetSetting(App.EXEName, "PreviewWinPos", "X")
        Left = GetSetting(App.EXEName, "PreviewWinPos", "Y")
        Width = GetSetting(App.EXEName, "PreviewWinPos", "W")
        Height = GetSetting(App.EXEName, "PreviewWinPos", "H")
        
        If Top >= Screen.Height Then Top = 0
        If Left >= Screen.Width Then Left = 0
        If Width < 2000 Then Width = 2000
        If Height < 2000 Then Height = 2000
    End If
    
    chkUseTarget.Value = Val(GetSetting(App.EXEName, "PreviewWinPos", "UseTarget", 0))
    SetColor GetSetting(App.EXEName, "PreviewWinPos", "BackColor", &HFFFFFF), cmdColor
    
    wBrowser.ZOrder 0
    
    FillCodepageCombo
    FixCtrls4Skin Me
    
    IsLoading = False
    
    tmrForceLoad.Enabled = True
    
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

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = 123 Then cmdRefresh_Click

End Sub

Private Sub Form_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

     If data.GetFormat(vbCFFiles) Then
        frmMain.LoadMenu data.Files(1)
        frmMain.ShowPreview
    End If

End Sub

Private Sub Form_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    If data.GetFormat(vbCFFiles) Then
        If Right$(data.Files(1), 3) = "dmb" Then
            Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectNone
        End If
    Else
        Effect = vbDropEffectNone
    End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    SaveSetting App.EXEName, "PreviewWinPos", "X", Top
    SaveSetting App.EXEName, "PreviewWinPos", "Y", Left
    SaveSetting App.EXEName, "PreviewWinPos", "W", Width
    SaveSetting App.EXEName, "PreviewWinPos", "H", Height
    SaveSetting App.EXEName, "PreviewWinPos", "BackColor", cmdColor.tag
    
    PreviewIsOn = False

End Sub

Private Sub Form_Resize()

    On Error Resume Next

    If WindowState <> vbMinimized Then
        If Width > 2000 And Height > 2000 Then
            With wBrowser
                .Move 60, 60, Width - 235, Height - 330 - sbSizer.Height - frameControls.Height
                frameControls.Move 60, .Top + .Height + 30, .Width, frameControls.Height
            End With
                        
            With frameControls
                cmdClose.Move .Width - (cmdClose.Width + 120)
                cmdRefresh.Move cmdClose.Left - (cmdRefresh.Width + 120)
                If IsSkinned Then Controls("dynpic1").Width = .Width - 60
            End With
        End If
        With cmbCharset
            If Width > .Left + 2565 + 300 Then
                .Width = 2565
            Else
                .Width = Width - .Left - 300
            End If
        End With
    End If
    
End Sub

Private Sub frameControls_OLEDragDrop(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

    If data.GetFormat(vbCFFiles) Then
        frmMain.LoadMenu data.Files(1)
        frmMain.ShowPreview
    End If

End Sub

Private Sub frameControls_OLEDragOver(data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)

    If data.GetFormat(vbCFFiles) Then
        If Right$(data.Files(1), 3) = "dmb" Then
            Effect = vbDropEffectCopy
        Else
            Effect = vbDropEffectNone
        End If
    Else
        Effect = vbDropEffectNone
    End If

End Sub

Private Sub tmrForceLoad_Timer()

    If tmrLoadPreview.Enabled Then Exit Sub
    tmrLoadPreview_Timer

End Sub

Private Sub tmrLoadPreview_Timer()

    tmrForceLoad.Enabled = False
    tmrLoadPreview.Enabled = False
    LoadPreview
    
End Sub

Private Sub wBrowser_NavigateComplete2(ByVal pDisp As Object, url As Variant)

    On Error Resume Next

    SetDocBackColor

    If url = "http:///" Or url = "about:blank" Then Exit Sub
    
    wBrowser.SetFocus

End Sub

Private Sub LocalizeUI()

    cmdRefresh.caption = GetLocalizedStr(955)
    cmdClose.caption = GetLocalizedStr(424)
    
    lblEncoding.caption = GetLocalizedStr(956)
    lblBackColor.caption = GetLocalizedStr(957)
    chkUseTarget.caption = GetLocalizedStr(958)
    
    If Preferences.language <> "eng" Then
        cmdRefresh.Width = SetCtrlWidth(cmdRefresh)
        cmdClose.Width = SetCtrlWidth(cmdClose)
        cmbCharset.Left = lblEncoding.Left + lblEncoding.Width + 30
    End If

End Sub
