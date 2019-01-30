Attribute VB_Name = "DMBGlobals"
Option Explicit

Global USER As String
Global COMPANY As String
Global DMBPSN As String
Global USERSN As String
Global ORDERNUMBER As String
Global DoSilentValidation As Boolean

Global Const CacheSignature = 1122

Global ResellerID As String
Global ResellerInfo() As String

Global InitDone As Boolean
Global IsDebug As Boolean
Global IsTBMapSel As Boolean

Global frmErrDlgIsVisible As Boolean

Global Const SupportedImageFiles = "Supported Image Files (*.jpg;*.gif;*.png)|*.jpg;*.gif;*.png;|All Files (*.*)|*.*"
Global Const SupportedCursorFiles = "Supported Cursor Files (*.cur;*.csr;*.ani)|*.cur;*.csr;*.ani|All Files (*.*)|*.*"
Global Const SupportedImageFilesFlash = "Supported Image Files (*.jpg;*.gif;*.png)|*.jpg;*.gif;*.png;|Macromedia Flash Movies (*.swf)|*.swf|All Files (*.*)|*.*"
Global Const SupportedHTMLDocs = "HTML Documents (*.htm;*.html;*.asp;*.aspx;*.ascx;*.php;*.php3;*.shtml;*.jsp;*.dwt;*.cfm;*.master)|*.htm;*.html;*.asp;*.aspx;*.ascx;*.php;*.php3;*.shtml;*.dwt;*.cfm;*.master|All Files (*.*)|*.*"
Global Const SupportedCSSDocs = "HyperText StyleSheets (*.css)|*.css|All Files (*.*)|*.*"
Global Const SupportedImportDocs = "hierMenus 4 (HM_Arrays.js)|HM_Arrays.js|hierMenus (hierArrays.js)|hierArrays.js|AllWebMenus (*.awm)|*.awm|All Files (*.*)|*.*"
Global Const SupportedAudioFiles = "Wav Files (*.wav)|*.wav|All Files (*.*)|*.*"

Public TemplateCommand As MenuCmd
Public TemplateGroup As MenuGrp
Public dmbClipboard As ClipboardDef
Public FramesInfo As FramesInfoDef
Public Sections() As Section
Public LocalizedStr() As String
Public engLocalizedStr() As String
Public NullFont As tFont
Public InMapMode As Boolean
Public DontRefreshMap As Boolean
Public KeepExpansions As Boolean
Public LastSelNode As String

Public SelSecProjects() As String
Public SelSecProjectsTitles() As String
Public Enum SecProjModeConstants
    spmcFromInstallMenus = 0
    spmcFromStdDlg = 1
    spmcUndefined = 999
End Enum
Public SecProjMode As SecProjModeConstants

Public Enum LinkVerifyModeConstants
    spmcManual = 0
    spmcAuto = 1
End Enum
Public LinkVerifyMode As LinkVerifyModeConstants

Public Enum ProjectPropertiesPageConstants
    pppcGeneral = 0
    pppcConfig = 1
    pppcGlobal = 2
    pppcAdvanced = 3
End Enum
Public ProjectPropertiesPage As ProjectPropertiesPageConstants

Public Enum TBEPageConstants
    tbepcGeneral = 0
    tbepcAppearance = 1
    tbepcPositioning = 2
    tbepcEffects = 3
    tbepcAdvanced = 4
End Enum
Public TBEPage As TBEPageConstants

Public Going2Upgrade As Boolean
Public dlFileName As String
Public AutoHotSpot As Boolean
Public SelColor As Long
Public SelColor_CanBeTransparent As Boolean
Public UsedColors() As Long
Public SimonFile As String
Public cSep As String
Public ff As Integer
Public HSCanceled As Boolean
Public SelImgName As String
Public AbortCompileDlg As Boolean
Public ImportProjectFileName As String
Public ppSelConfig As Integer
Public IsReplacing As Boolean
Public PreviewIsOn As Boolean
Public MenusFrame As Integer
Public NagScreenIsVisible As Boolean
Public IsFPAddIn As Boolean
Public curOffsetStr As String

Public nwdPar As String

#If ISCOMP = 0 Then
Public TipsSys As CTips
Public wbLivePreview As WebBrowser
Public LivePreviewCharset As String
#End If

Public Enum PreviewModeConstants
    pmcNormal = 0
    pmcSitemap = 1
End Enum
Public PreviewMode As PreviewModeConstants

Public Type SelImageDef
    Picture As IPictureDisp
    FileName As String
    IsResource As Boolean
    IsValid As Boolean
    SupportsFlash As Boolean
    LimitToCursors As Boolean
End Type
Public SelImage As SelImageDef

Public Type UndoState
    FileName As String
    Description As String
End Type
Public UndoStates() As UndoState
Public CurState As Integer

Public Enum SelRedoUndoConstants
    [sUndo]
    [sRedo]
    [sCancel]
End Enum
Public SelRedoUndo As SelRedoUndoConstants
Public SelRedoUndoCount As Integer

Public Type SelFontDef
    Name As String
    Italic As Boolean
    Bold As Boolean
    Underline As Boolean
    Size As Integer
    IsValid As Boolean
    IsSubst As Boolean
    Shadow As tFontShadow
End Type
Public SelFont As SelFontDef

Private OriginalConfig As Integer

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Public Const LVM_FIRST = &H1000
Public Const LVM_SETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 54
Public Const LVM_GETEXTENDEDLISTVIEWSTYLE = LVM_FIRST + 55

Public Const LVS_EX_FULLROWSELECT = &H20
Public Const LVSCW_AUTOSIZE = -1
Public Const LVSCW_AUTOSIZE_USEHEADER = -2
Public Const LVM_SETCOLUMNWIDTH As Long = &H101E
Public Const HDS_BUTTONS = &H2
Public Const LVM_GETHEADER = (LVM_FIRST + 31)
Public Const GWL_STYLE = (-16)

Public Const CB_GETLBTEXTLEN = &H149
Public Const CB_SHOWDROPDOWN = &H14F
Public Const CB_GETDROPPEDWIDTH = &H15F
Public Const CB_SETDROPPEDWIDTH = &H160

Private Const MF_BYPOSITION = &H400
Private Const MF_REMOVE = &H1000

Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long

Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    bottom As Long
End Type
Public Const BDR_INNER = &HC
Public Const BDR_OUTER = &H3
Public Const BDR_RAISED = &H5
Public Const BDR_RAISEDINNER = &H4
Public Const BDR_RAISEDOUTER = &H1
Public Const BDR_SUNKEN = &HA
Public Const BDR_SUNKENINNER = &H8
Public Const BDR_SUNKENOUTER = &H2
Public Const BF_ADJUST = &H2000   ' Calculate the space left over.
Public Const BF_BOTTOM = &H8
Public Const BF_LEFT = &H1
Public Const BF_RIGHT = &H4
Public Const BF_TOP = &H2
Public Const BF_BOTTOMLEFT = (BF_BOTTOM Or BF_LEFT)
Public Const BF_BOTTOMRIGHT = (BF_BOTTOM Or BF_RIGHT)
Public Const BF_TOPLEFT = (BF_TOP Or BF_LEFT)
Public Const BF_TOPRIGHT = (BF_TOP Or BF_RIGHT)
Public Const BF_DIAGONAL = &H10
Public Const BF_DIAGONAL_ENDBOTTOMLEFT = (BF_DIAGONAL Or BF_BOTTOM Or BF_LEFT)
Public Const BF_DIAGONAL_ENDBOTTOMRIGHT = (BF_DIAGONAL Or BF_BOTTOM Or BF_RIGHT)
Public Const BF_DIAGONAL_ENDTOPLEFT = (BF_DIAGONAL Or BF_TOP Or BF_LEFT)
Public Const BF_DIAGONAL_ENDTOPRIGHT = (BF_DIAGONAL Or BF_TOP Or BF_RIGHT)
Public Const BF_FLAT = &H4000     ' For flat rather than 3-D borders.
Public Const BF_MIDDLE = &H800    ' Fill in the middle.
Public Const BF_MONO = &H8000     ' For monochrome borders.
Public Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Const BF_SOFT = &H1000     ' Use for softer buttons.
Public Declare Function DrawEdge Lib "user32" (ByVal hDc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Public Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDc As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

'Private Type MIMECPINFO
'    dwFlags As Long
'    uiCodePage As Integer
'    uiFamilyCodePage As Integer
'    wszDescription As String * 255
'    wszWebCharSet As String * 255
'    wszHeaderCharset As String * 255
'    wszBodyCharset As String * 255
'    wszFixedWidthFont As String * 255
'    wszProportionalFonts As String * 255
'End Type
Private Declare Function LcidToRfc1766 Lib "mlang.dll" Alias "LcidToRfc1766A" (ByVal uiCodePage As Long, ByRef pszRfc1766 As String, ByRef nChar As Integer) As Long

Public Type CodePagesDef
    CodePage As String
    WebCharset As String
    Description As String
End Type
Public cs() As CodePagesDef

#If STANDALONE = 0 Then

Public ftpUserName As String
Public ftpPassword As String

Public Function GetAppTypeName() As String

    Select Case GetAppType
        Case "DEV"
            GetAppTypeName = "<developers' edition>"
        Case "STD"
            GetAppTypeName = "Standard Edition"
        Case "LIT"
            GetAppTypeName = "LITE"
    End Select

End Function

Public Function GetAppType() As String

    #If LITE = 0 Then
        #If DEVVER = 1 Then
            GetAppType = "DEV"
        #Else
            GetAppType = "STD"
        #End If
    #Else
        GetAppType = "LIT"
    #End If

End Function

#If ISCOMP = 0 Then

Public Sub LoadUnicodeTool()

    On Error Resume Next
    RunShellExecute "open", "dmbUnicodeCaption.exe", "", Long2Short(AppPath), 1

End Sub

Public Sub PopulateBorderStyleCombo(cmb As ComboBox)

    With cmb
        .Clear
        .AddItem GetLocalizedStr(110)
        .AddItem GetLocalizedStr(430)
        .AddItem GetLocalizedStr(431)
        .AddItem GetLocalizedStr(670)
        .AddItem GetLocalizedStr(671)
        .AddItem "Dotted"
        .AddItem "Dashed"
    End With

End Sub

Public Function GetSysCharsets() As CodePagesDef()

    Dim i As Long
    Dim j As Long
    Dim CodePage As String
    Dim WebCharset As String
    Dim Desc As String
    Static cs() As CodePagesDef
    Dim tmp As CodePagesDef
    
    On Error Resume Next
    
    ReDim cs(0)
    
    i = 0
    Do
        CodePage = EnumSubKeys(HKEY_CLASSES_ROOT, "MIME\Database\Codepage", i)
        If LenB(CodePage) <> 0 Then
            WebCharset = QueryValue(HKEY_CLASSES_ROOT, "MIME\Database\Codepage\" + CodePage, "WebCharset")
            If LenB(WebCharset) = 0 Then
                WebCharset = QueryValue(HKEY_CLASSES_ROOT, "MIME\Database\Codepage\" + CodePage, "BodyCharset")
                Desc = QueryValue(HKEY_CLASSES_ROOT, "MIME\Database\Codepage\" + CodePage, "Description")
            Else
                Desc = QueryValue(HKEY_CLASSES_ROOT, "MIME\Database\Codepage\" + CodePage, "Description")
            End If
            
            If InStr(Desc, "@") > 0 Then
                ' We must be under Vista...
                Desc = String(255, vbNullChar)
                Dim n As Integer
                If LcidToRfc1766(CLng(CodePage), Desc, n) = 0 Then
                    Desc = Mid(Desc, 1, n)
                Else
                    Desc = WebCharset
                End If
            End If
            
            If LenB(WebCharset) <> 0 Then
                ReDim Preserve cs(UBound(cs) + 1)
                With cs(UBound(cs))
                    .CodePage = CodePage
                    .Description = Desc
                    .WebCharset = WebCharset
                End With
            End If
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    
    For i = 1 To UBound(cs)
        For j = 1 To UBound(cs)
            If cs(j).Description > cs(i).Description Then
                tmp = cs(j)
                cs(j) = cs(i)
                cs(i) = tmp
            End If
        Next j
    Next i
    
    GetSysCharsets = cs

End Function
#End If

Public Function strGetSupportedHTMLDocs() As String

    strGetSupportedHTMLDocs = Split(Replace(UCase(SupportedHTMLDocs), "*.", vbNullString), "|")(1)

End Function

Public Function Hex2Dec(h As String) As Long

    On Error Resume Next

    If LenB(h) = 0 Then h = 0
    Hex2Dec = CLng("&H" & h)

End Function

Public Function SaveProject(IsState As Boolean) As Boolean

    Dim i As Integer
    Dim ff As Integer
    Dim sStr As String
    Dim k As Integer
    Dim c As Integer
    Dim g As Integer
    Dim t As Integer
    
    On Error GoTo ReportError
    
    If Not IsState Then
        If LCase(GetFileExtension(Project.FileName)) <> "dmb" Then
            Project.FileName = GetFilePath(Project.FileName) + GetFileName(Project.FileName, True) + ".dmb"
        End If
        
        If Not RemoveReadOnly(Project.FileName) Then
            MsgBox "The project could not be saved because its marked Read Only", vbCritical + vbOKOnly, "Error Saving Project"
            SaveProject = False
            Exit Function
        End If
    End If
    
    ff = FreeFile
    Open Project.FileName For Output As ff

        With Project
            sStr = .Name + _
                    cSep & Abs(.SEOTweak) & _
                    cSep & .UserConfigs(0).CompiledPath & _
                    cSep & .UserConfigs(0).HotSpotEditor.HotSpotsFile & _
                    cSep & .FX & _
                    cSep & .NodeExpStatus & _
                    cSep & .ToolBar.BorderColor & _
                    cSep & .CodeOptimization & _
                    cSep & Abs(.UserConfigs(0).HotSpotEditor.MakeBackup) & _
                    cSep & GetCurProjectVersion & _
                    cSep & Abs(.UserConfigs(0).Frames.UseFrames) & _
                    cSep & .AddIn.Name & _
                    cSep & .UserConfigs(0).Frames.FramesFile & _
                    cSep & .UserConfigs(0).RootWeb & _
                    cSep & .SelChangeDelay & _
                    cSep & GenExpHTMLString & _
                    cSep & 0 & cSep & 0 & cSep & 0 & _
                    cSep & 0 & cSep & 0 & cSep & 0 & _
                    cSep & 0 & cSep & 0 & cSep & 0 & _
                    cSep & 0 & cSep & 0 & cSep & 0 & _
                    cSep & 0
            'sStr = sStr + _
                    cSep & .FTP.FTPAddress & _
                    cSep & .FTP.UserName & _
                    cSep & "" & _
                    cSep & .FTP.ProxyAddress & _
                    cSep & .FTP.ProxyPort & _
                    cSep & .FTP.RemoteInfo4FTP & _
                    cSep & .UserConfigs(0).ImagesPath & _
                    cSep & .JSFileName & _
                    cSep & UBound(.UserConfigs) & _
                    cSep & .DefaultConfig
            sStr = sStr + _
                    cSep & .UserConfigs(0).FTP & _
                    cSep & vbNullString & _
                    cSep & Abs(.UseGZIP) & _
                    cSep & .BlinkSpeed & _
                    cSep & .BlinkEffect & _
                    cSep & .CustomOffsets & _
                    cSep & .UserConfigs(0).ImagesPath & _
                    cSep & .JSFileName & _
                    cSep & UBound(.UserConfigs) & _
                    cSep & .DefaultConfig
                    
            For i = 1 To UBound(.UserConfigs)
                With .UserConfigs(i)
                    sStr = sStr + _
                        cSep & .Name & _
                        cSep & .Description & _
                        cSep & .CompiledPath & _
                        cSep & .Type & _
                        cSep & .RootWeb & _
                        cSep & .ImagesPath & _
                        cSep & Abs(.OptmizePaths) & _
                        cSep & .FTP & _
                        cSep & .Frames.FramesFile & _
                        cSep & vbNullString & _
                        cSep & Abs(.Frames.UseFrames) & _
                        cSep & .HotSpotEditor.HotSpotsFile & _
                        cSep & Abs(.HotSpotEditor.MakeBackup) & _
                        cSep & .LocalInfo4RemoteConfig
                End With
            Next i
            sStr = sStr & cSep & .UnfoldingSound.onmouseover
            sStr = sStr & cSep & .MenusOffset.RootMenusX
            sStr = sStr & cSep & .MenusOffset.RootMenusY
            sStr = sStr & cSep & .MenusOffset.SubMenusX
            sStr = sStr & cSep & .MenusOffset.SubMenusY
            
            sStr = sStr & cSep & Abs(.LotusDominoSupport)
            
            #If DEVVER = 1 Then
                sStr = sStr & cSep & Abs(.GenDynAPI)
            #Else
                sStr = sStr & cSep & 0
            #End If
            
            sStr = sStr & cSep & Abs(.CompileIECode)
            sStr = sStr & cSep & Abs(.CompileNSCode)
            sStr = sStr & cSep & Abs(.CompilehRefFile)
            
            sStr = sStr & cSep & Join(.SecondaryProjects, "|")
            
            sStr = sStr & cSep & .FontSubstitutions
            sStr = sStr & cSep & Abs(.DoFormsTweak)
            
            sStr = sStr & cSep & .StatusTextDisplay
            sStr = sStr & cSep & Abs(.KeyboardSupport)
            sStr = sStr & cSep & Abs(.RemoveImageAutoPosCode)
            sStr = sStr & cSep & .RootMenusDelay
            
            sStr = sStr & cSep & .AnimSpeed
            sStr = sStr & cSep & .HideDelay
            sStr = sStr & cSep & .SubMenusDelay
            
            sStr = sStr & cSep & Abs(.DWSupport)
            sStr = sStr & cSep & Abs(.NS4ClipBug)
            
            sStr = sStr & cSep & ctcCDROM
            
            sStr = sStr & cSep & Abs(.AutoSelFunction)
            sStr = sStr & cSep & Abs(.ImageReadySupport)
            
            sStr = sStr & cSep & .DXFilter
            
            With .AutoScroll
                sStr = sStr + cSep & .maxHeight
                sStr = sStr + cSep & .nColor
                sStr = sStr + cSep & .hColor
                sStr = sStr + cSep & .DnImage.NormalImage
                sStr = sStr + cSep & .DnImage.HoverImage
                sStr = sStr + cSep & .DnImage.w
                sStr = sStr + cSep & .DnImage.h
                sStr = sStr + cSep & .UpImage.NormalImage
                sStr = sStr + cSep & .UpImage.HoverImage
                sStr = sStr + cSep & .margin
                sStr = sStr + cSep & Abs(.onmouseover)
                sStr = sStr + cSep & .FXhColor
                sStr = sStr + cSep & .FXnColor
                sStr = sStr + cSep & .FXNormal
                sStr = sStr + cSep & .FXOver
                sStr = sStr + cSep & .FXSize
            End With
            
            For i = 1 To 50 - 16
                sStr = sStr & cSep & vbNullString
            Next i
            
            sStr = sStr & cSep & Abs(UBound(Project.Toolbars))
            For i = 1 To UBound(.Toolbars)
                With .Toolbars(i)
                    sStr = sStr & cSep & .Alignment
                    sStr = sStr & cSep & .BackColor
                    sStr = sStr & cSep & .bOrder
                    sStr = sStr & cSep & .BorderColor
                    sStr = sStr & cSep & .ContentsMarginH
                    sStr = sStr & cSep & .ContentsMarginV
                    sStr = sStr & cSep & .CustX
                    sStr = sStr & cSep & .CustY
                    sStr = sStr & cSep & Abs(.FollowHScroll)
                    If .FollowVScroll = False And .SmartScrolling = False Then sStr = sStr & cSep & 0
                    If .FollowVScroll = True And .SmartScrolling = False Then sStr = sStr & cSep & 1
                    If .FollowVScroll = True And .SmartScrolling = True Then sStr = sStr & cSep & 2
                    If .FollowVScroll = False And .SmartScrolling = True Then sStr = sStr & cSep & 3
                    sStr = sStr & cSep & Join(.Groups, "|")
                    sStr = sStr & cSep & .Height
                    sStr = sStr & cSep & .Image
                    sStr = sStr & cSep & Abs(.JustifyHotSpots)
                    sStr = sStr & cSep & .OffsetH
                    sStr = sStr & cSep & .OffsetV
                    sStr = sStr & cSep & .Name
                    sStr = sStr & cSep & .Separation
                    sStr = sStr & cSep & .Spanning
                    sStr = sStr & cSep & .Style
                    sStr = sStr & cSep & .Width
                    sStr = sStr & cSep & .AttachTo + "|" & .AttachToAlignment
                    sStr = sStr & cSep & .Condition
                    sStr = sStr & cSep & .BorderStyle
                    sStr = sStr & cSep & .DropShadowColor
                    sStr = sStr & cSep & .DropShadowSize
                    sStr = sStr & cSep & .Transparency
                    sStr = sStr & cSep & Abs(.IsTemplate)
                    sStr = sStr & cSep & Abs(.AttachToAutoResize)
                    sStr = sStr & cSep & Abs(.Compile)
                    
                    sStr = sStr & cSep & .Radius.TopLeft
                    sStr = sStr & cSep & .Radius.TopRight
                    sStr = sStr & cSep & .Radius.BottomLeft
                    sStr = sStr & cSep & .Radius.BottomRight
                End With
            Next i
            
            Print #ff, sStr
        End With
        
        t = UBound(MenuGrps) + UBound(MenuCmds)
        For g = 1 To UBound(MenuGrps)
            If Not IsState Then
                k = k + 1: FloodPanel.Value = k / t * 100
            End If
            Print #ff, "[G]" + GetGrpParams(MenuGrps(g))
            For c = 1 To UBound(MenuCmds)
                With MenuCmds(c)
                    If .parent = g Then
                        If Not IsState Then
                            k = k + 1: FloodPanel.Value = k / t * 100
                        End If
                        Print #ff, "[C]  " + GetCmdParams(MenuCmds(c))
                    End If
                End With
            Next c
        Next g
    
    SaveProject = True
ExitSub:
    Close #ff
    
    FloodPanel.Value = 0
    
    Exit Function
    
ReportError:
    SaveProject = False
    MsgBox "An error has occured while saving the project" + vbCrLf + "Error (" & Err.number & "): " & Err.Description, vbCritical + vbOKOnly, "Error Saving Project"
    GoTo ExitSub
    
End Function

Public Function SubMenuOf(g As Integer) As Integer

    Dim i As Integer
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i).Actions
            If .onclick.Type = atcCascade And .onclick.TargetMenu = g Then
                SubMenuOf = i
                Exit Function
            End If
            If .onmouseover.Type = atcCascade And .onmouseover.TargetMenu = g Then
                SubMenuOf = i
                Exit Function
            End If
            If .OnDoubleClick.Type = atcCascade And .OnDoubleClick.TargetMenu = g Then
                SubMenuOf = i
                Exit Function
            End If
        End With
    Next i

End Function

Public Function GenExpHTMLPref(c As String, ProjectName As String, ProjectFileName As String) As ExportHTMLDef

    Dim p() As String
    
    On Error Resume Next
    
    If InStr(c, "|") Then
        p = Split(c, "|")
        With GenExpHTMLPref
            .CollapsedImage = p(0)
            .CommandClass = p(1)
            .CreateTree = -Val(p(2))
            .CSSFile = p(3)
            .Description = p(4)
            .ExpandedImage = p(5)
            .GroupClass = p(6)
            .HTMLFileName = p(7)
            .IconHeight = Val(p(8))
            .IconWidth = Val(p(9))
            .Identation = Val(p(10))
            .ImagesPath = p(11)
            .NormalImage = p(12)
            .Style = Val(p(13))
            .Title = p(14)
            .ExpItemsHaveLinks = -Val(p(15))
            .SingleSelect = -Val(p(16))
            .IncludeExpCol = -Val(p(17))
            .ExpAllStr = p(18)
            .ColAllStr = p(19)
            .ExpColPlacement = p(20)
            If UBound(p) > 20 Then
                .XHTMLCompliant = -Val(p(21))
            End If
        End With
    Else
        With GenExpHTMLPref
            .CollapsedImage = AppPath + "exhtml\c.gif"
            .CommandClass = vbNullString
            .CreateTree = False
            .CSSFile = vbNullString
            #If ISCOMP = 0 Or STANDALONE = 1 Then
            .Description = "<small>DHTML Menu Builder " + DMBVersion + "</small>"
            #Else
            .Description = "<small>DHTML Menu Builder Wizard</small>"
            #End If
            .ExpandedImage = AppPath + "exhtml\o.gif"
            .GroupClass = vbNullString
            .HTMLFileName = GetFilePath(ProjectFileName) + GetFileName(ProjectFileName) + ".htm"
            .IconHeight = 16
            .IconWidth = 16
            .Identation = 20
            .ImagesPath = "%%PROJECT%%"
            .NormalImage = AppPath + "exhtml\s.gif"
            .Style = ascProject
            .Title = ProjectName
            .ExpItemsHaveLinks = False
            .SingleSelect = False
            .IncludeExpCol = False
            .ExpAllStr = "Expand All"
            .ColAllStr = "Collapse All"
            .ExpColPlacement = ecpcBottom
            .XHTMLCompliant = False
        End With
    End If

End Function

Private Function GenExpHTMLString() As String

    With Project.ExportHTMLParams
        GenExpHTMLString = .CollapsedImage & "|" & _
                           .CommandClass & "|" & _
                           Abs(.CreateTree) & "|" & _
                           .CSSFile & "|" & _
                           .Description & "|" & _
                           .ExpandedImage & "|" & _
                           .GroupClass & "|" & _
                           .HTMLFileName & "|" & _
                           .IconHeight & "|" & _
                           .IconWidth & "|" & _
                           .Identation & "|" & _
                           .ImagesPath & "|" & _
                           .NormalImage & "|" & _
                           .Style & "|" & _
                           .Title & "|" & _
                           Abs(.ExpItemsHaveLinks) & "|" & _
                           Abs(.SingleSelect) & "|" & _
                           Abs(.IncludeExpCol) & "|" & _
                           .ExpAllStr & "|" & _
                           .ColAllStr & "|" & _
                           .ExpColPlacement & "|" & _
                           Abs(.XHTMLCompliant)
    End With

End Function

#End If

Public Function AddTrailingSlash(ByVal Path As String, Slash As String) As String

    If Right$(Path, 1) <> Slash And LenB(Path) <> 0 Then Path = Path + Slash
    AddTrailingSlash = Path

End Function

Public Function GetTEMPPath() As String

    Dim TempPath As String

    TempPath = Environ("TEMP")
    If LenB(TempPath) = 0 Then TempPath = Environ("TMP")
    If LenB(TempPath) = 0 Then TempPath = AppPath
    If InStr(TempPath, ";") > 0 Then TempPath = Split(TempPath, ";")(0)
    
    TempPath = AddTrailingSlash(TempPath, "\")
    
    GetTEMPPath = TempPath

End Function

Public Sub TileImage(ImagePath As String, trg As PictureBox, Optional PreventTile As Boolean)

    Dim x As Long
    Dim y As Long
    Dim Src As PictureBox
    Dim f As Integer
    Dim IsAutoRedraw As Boolean
    
    Set Src = frmMain.picRsc
    
    f = 1
    If trg.ScaleMode = vbPixels Then f = 15
    
    trg.Picture = LoadPicture()
    IsAutoRedraw = trg.AutoRedraw
    trg.AutoRedraw = True
    If ImagePath <> "picRsc" Then
        If IsFlash(ImagePath) Then
            Src.Picture = LoadResPicture(200, vbResIcon)
        Else
            'If IsANI(ImagePath) Then
            '    src.Picture = LoadResPicture(201, vbResIcon)
            'Else
                Src.Picture = LoadPictureRes(ImagePath)
            'End If
        End If
    End If
    If Src <> 0 Then
        If PreventTile Then
            trg.PaintPicture Src, 0, 0
        Else
            For x = 0 To trg.Width Step Src.Width / f
                For y = 0 To trg.Height Step Src.Height / f
                    trg.PaintPicture Src, x, y
                    'BitBlt trg.hDC, x, Y, src.Width, src.Height, src.hDC, 0, 0, vbSrcCopy
                Next y
            Next x
        End If
    End If
    trg.AutoRedraw = IsAutoRedraw
    
End Sub

Public Sub SetTemplateDefaults()

    nwdPar = "NewWindow" + _
            cSep + "0" + _
            cSep + "0" + _
            cSep + "400" + _
            cSep + "400" + _
            cSep + "1" + _
            cSep + "0" + _
            cSep + "1" + _
            cSep + "0" + _
            cSep + "1" + _
            cSep + "1" + _
            cSep + "1" + _
            cSep + "1" + _
            cSep + "1" + _
            cSep + "1"
    
    With TemplateCommand
        .Name = vbNullString
        .caption = vbNullString
        .WinStatus = vbNullString
        .nBackColor = &HE0E0E0
        .nTextColor = &H0
        .hBackColor = &H947A4E
        .hTextColor = &HFFFFFF
        .iCursor.cType = iccDefault
        .iCursor.CFile = vbNullString
        .parent = 0
        .Trigger = ByClicking
        .NormalFont.FontName = "Tahoma"
        .NormalFont.FontSize = 11
        .NormalFont.FontBold = False
        .NormalFont.FontItalic = False
        .NormalFont.FontUnderline = False
        .HoverFont.FontName = "Tahoma"
        .HoverFont.FontSize = 11
        .HoverFont.FontBold = False
        .HoverFont.FontItalic = False
        .HoverFont.FontUnderline = False
        .Actions.onclick.TargetFrame = "_self"
        .Actions.onclick.Type = atcURL
        .Actions.onclick.url = vbNullString
        .Actions.onclick.WindowOpenParams = nwdPar
        .Actions.onmouseover.TargetFrame = "_self"
        .Actions.onmouseover.Type = atcNone
        .Actions.onmouseover.url = vbNullString
        .Actions.onmouseover.WindowOpenParams = nwdPar
        .Actions.OnDoubleClick.TargetFrame = "_self"
        .Actions.OnDoubleClick.Type = atcNone
        .Actions.OnDoubleClick.url = vbNullString
        .Actions.OnDoubleClick.WindowOpenParams = nwdPar
        .Alignment = tacLeft
        .disabled = False
        .LeftImage.h = 0
        .LeftImage.w = 0
        .LeftImage.HoverImage = vbNullString
        .LeftImage.NormalImage = vbNullString
        .RightImage.h = 0
        .RightImage.w = 0
        .RightImage.HoverImage = vbNullString
        .RightImage.NormalImage = vbNullString
        .BackImage.NormalImage = vbNullString
        .BackImage.HoverImage = vbNullString
        .BackImage.Tile = True
        .BackImage.AllowCrop = True
        .Sound.onmouseover = vbNullString
        .Sound.onclick = vbNullString
        .SeparatorPercent = 80
        
        .CmdsFXNormal = cfxcNone
        .CmdsFXOver = cfxcNone
        .CmdsFXSize = 1
        .CmdsMarginX = 6
        .CmdsMarginY = 3
        .CmdsFXnColor = -2
        .CmdsFXhColor = -2
        
        With .Radius
            .TopLeft = 0
            .TopRight = 0
            .BottomLeft = 0
            .BottomRight = 0
        End With
        
        .xData = vbNullString
        
        .Compile = True
    End With
    
    With TemplateGroup
        .Name = vbNullString
        .bColor = &H808080
        .Corners.leftCorner = &H808080
        .Corners.topCorner = &H808080
        .Corners.rightCorner = &H808080
        .Corners.bottomCorner = &H808080
        .Actions.onclick.Type = atcNone
        .Actions.onclick.TargetMenu = 0
        .Actions.onclick.url = vbNullString
        .Actions.onclick.TargetFrame = "_self"
        .Actions.onclick.WindowOpenParams = nwdPar
        .Actions.onmouseover.Type = atcNone
        .Actions.onmouseover.TargetMenu = 0
        .Actions.onmouseover.url = vbNullString
        .Actions.onmouseover.TargetFrame = "_self"
        .Actions.onmouseover.WindowOpenParams = nwdPar
        .Actions.OnDoubleClick.Type = atcNone
        .Actions.OnDoubleClick.TargetMenu = 0
        .Actions.OnDoubleClick.url = vbNullString
        .Actions.OnDoubleClick.TargetFrame = "_self"
        .Actions.OnDoubleClick.WindowOpenParams = nwdPar
        .frameBorder = 1
        .DefNormalFont.FontBold = False
        .DefNormalFont.FontItalic = False
        .DefNormalFont.FontName = "Tahoma"
        .DefNormalFont.FontSize = 11
        .DefNormalFont.FontUnderline = False
        .DefHoverFont = .DefNormalFont
        .x = 0
        .y = 0
        .Leading = 1
        .Image = vbNullString
        .Alignment = gacBottomLeft
        .CmdsFXNormal = cfxcNone
        .CmdsFXOver = cfxcNone
        .CmdsFXSize = 1
        .CmdsMarginX = 6
        .CmdsMarginY = 3
        .CmdsFXnColor = -2
        .CmdsFXhColor = -2
        .DropShadowSize = 0
        .DropShadowColor = &H999999
        .Transparency = 0
        .iCursor.cType = iccDefault
        .iCursor.CFile = vbNullString
        .fWidth = 0
        .CaptionAlignment = tacCenter
        .hBackColor = &H947A4E
        .hTextColor = &HFFFFFF
        .nBackColor = &HE0E0E0
        .nTextColor = &H0
        .disabled = False
        .caption = vbNullString
        .WinStatus = vbNullString
               
        .ContentsMarginH = 0
        .ContentsMarginV = 0
        
        .HSImage = vbNullString
        
        .IsContext = False
        
        .Sound.onmouseover = vbNullString
        .Sound.onclick = vbNullString
        
        .AlignmentStyle = ascVertical
        
        .xData = vbNullString
        
        .CornersImages.gcTopLeft = vbNullString
        .CornersImages.gcTopCenter = vbNullString
        .CornersImages.gcTopRight = vbNullString
        .CornersImages.gcLeft = vbNullString
        .CornersImages.gcRight = vbNullString
        .CornersImages.gcBottomLeft = vbNullString
        .CornersImages.gcBottomCenter = vbNullString
        .CornersImages.gcBottomRight = vbNullString
        
        .BorderStyle = cfxcNone
        
        With .scrolling
            .UpImage.NormalImage = AppPath + "exhtml\aup_b.gif"
            .UpImage.HoverImage = AppPath + "exhtml\aup_w.gif"
            .DnImage.NormalImage = AppPath + "exhtml\adn_b.gif"
            .DnImage.HoverImage = AppPath + "exhtml\adn_w.gif"
            
            .FXhColor = &H0
            .FXnColor = &H0
            .FXNormal = cfxcNone
            .FXOver = cfxcNone
            .FXSize = 1
            .hColor = &H202080
            .margin = 4
            .nColor = &H808080
            .onmouseover = True
        End With
        
        .tbiLeftImage.NormalImage = vbNullString
        .tbiLeftImage.HoverImage = vbNullString
        .tbiLeftImage.w = 0
        .tbiLeftImage.h = 0
        
        .tbiRightImage.NormalImage = vbNullString
        .tbiRightImage.HoverImage = vbNullString
        .tbiRightImage.w = 0
        .tbiRightImage.h = 0
        
        .tbiBackImage.NormalImage = vbNullString
        .tbiBackImage.HoverImage = vbNullString
        .tbiBackImage.Tile = True
        .tbiBackImage.AllowCrop = True
        
        With .Radius
            .TopLeft = 0
            .TopRight = 0
            .BottomLeft = 0
            .BottomRight = 0
        End With
        
        .Compile = True
    End With

End Sub

#If STANDALONE = 0 Then

#If ISCOMP = 0 Then
Public Sub ResizeComboList(c As ComboBox)

    Dim i As Integer
    Dim w As Long
    Dim tw As Integer

    For i = 0 To c.ListCount - 1
        tw = GetTextSize(c.List(i), , c, vbPixels)(1)
        If tw > w Then w = tw
    Next i

    SendMessage c.hwnd, CB_SETDROPPEDWIDTH, w + 20, ByVal 0

End Sub
#End If

Public Function SetCtrlWidth(ctrl As Control, Optional CanReduce As Boolean = False) As Long

    Dim w As Integer
    
    w = GetTextSize(ctrl.caption, , ctrl)(1)
    If w + 150 > ctrl.Width Or CanReduce Then
        SetCtrlWidth = w + 150
    Else
        SetCtrlWidth = ctrl.Width
    End If

End Function

Public Function GetID(Optional nItm As Node) As Integer
    
    Dim i As Integer
    
    On Error Resume Next

    If nItm Is Nothing Then Set nItm = frmMain.tvMenus.SelectedItem
    If LenB(nItm.key) = 0 Then
        For i = 1 To UBound(MenuGrps)
            If MenuGrps(i).Name = nItm.Text Then
                nItm.key = "G" & i
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(MenuCmds)
            If MenuCmds(i).Name = nItm.Text Then
                nItm.key = "C" & i
                Exit For
            End If
        Next i
    End If
    If Not nItm Is Nothing Then GetID = Mid$(nItm.key, 2)

End Function

Public Sub CenterForm(frm As Form)
    
    If frm.Name = "frmMain" Or frm.Name = "frmNag" Then
        With Screen
            frm.Move (.Width - frm.Width) \ 2, (.Height - frm.Height) \ 2
        End With
    Else
        With frmMain
            frm.Move .Left + (.Width - frm.Width) \ 2, .Top + (.Height - frm.Height) \ 2
        End With
    End If

End Sub

Public Function IsGroup(key As String) As Boolean

    IsGroup = Left$(key, 1) = "G"

End Function

Public Function IsCommand(key As String) As Boolean

    IsCommand = Left$(key, 1) = "C"

End Function

Public Function IsTemplate(key As String) As Boolean

    If InMapMode And IsTBMapSel Then
        #If ISCOMP = 0 Then
        IsTemplate = Project.Toolbars(ToolbarIndexByKey(frmMain.tvMapView.SelectedItem.key)).IsTemplate
        #End If
    Else
        If IsGroup(key) Then
            IsTemplate = MenuGrps(Val(Mid$(key, 2))).IsTemplate
        Else
            IsTemplate = MenuGrps(MenuCmds(Val(Mid$(key, 2))).parent).IsTemplate
        End If
    End If
    
End Function

Public Function IsSeparator(key As String) As Boolean

    IsSeparator = Left$(key, 1) = "S"

End Function

Public Function GenLicense() As String

    Dim txtStr As String
    Dim i As Integer
    Dim acc As Double
    Dim acc1 As Double
    Dim acc2 As Double
    Dim h As String
    
    txtStr = USER + COMPANY
    
    For i = 1 To Len(txtStr) * 2
        If i <= Len(txtStr) Then
            acc = acc + Asc(Mid$(txtStr, i, 1)) + i
        Else
            acc = acc + Asc(Mid$(txtStr, i - Len(txtStr), 1)) ^ 2
        End If
    Next i
    
    For i = Len(txtStr) To 1 Step -1
        acc1 = acc1 + (Asc(Mid$(txtStr, i, 1)) + Asc(Mid$(txtStr, Len(txtStr) - i + 1, 1)) Xor i)
    Next i
    
    For i = 1 To Len(txtStr)
        acc2 = acc2 + Asc(Mid$(txtStr, i, 1)) * i
    Next i
    
    h = Hex(acc2)
    If Len(h) \ 2 <> Len(h) / 2 Then h = "0" + h
    
    If acc \ 2 = acc / 2 Then
        GenLicense = h + "-" + Format(acc, "000000") + "-XFX-" + Format(acc1, "0000")
    Else
        GenLicense = Format(acc, "0000") + "-TWB-" + Format(acc1, "0000") + "-" + h
    End If

End Function

Public Function GetProjectProperties(FileName As String, Optional parseToolbars As Boolean = True) As ProjectDef

    Dim sStr As String
    Dim i As Long
    Dim j As Long
    Dim tcp As Integer
    Dim IsOld As Boolean
    Dim IsPre41 As Boolean
    Dim IsPre42 As Boolean
    Dim IsPre45005 As Boolean
    Dim IsPre49011 As Boolean
    Dim IsPre49018 As Boolean
    Dim IsPre42030 As Boolean
    
    On Error Resume Next
    
    ff = FreeFile
    Open FileName For Binary As #ff
    Line Input #ff, sStr
    
    'For compatibility with older projects using the chr(255) as a delimiter
    For i = Len(sStr) To 2 Step -1
        If Mid(sStr, i, 1) = Chr(255) And Mid(sStr, i - 1, 1) <> Chr(255) Then
            MsgBox "The Project you have selected was created using an older version of DHTML Menu Builder." + vbCrLf + _
                   "If you save this project, will be saved using the new Projects Format which is incompatible with previous versions.", vbInformation + vbOKOnly, "Warning"
            cSep = Chr(255)
            Exit For
        End If
        If Mid(sStr, i, 1) = Chr(255) And Mid(sStr, i - 1, 1) = Chr(255) Then
            cSep = Chr(255) + Chr(255)
            Exit For
        End If
    Next i
    
    With GetProjectProperties
        ReDim .UserConfigs(0)
        
        'For compatibility with 2.x projects
        .UserConfigs(0).Name = "Local"
        .UserConfigs(0).Description = GetLocalizedStr(282)
        .UserConfigs(0).Type = ctcLocal
        .DefaultConfig = 0
        
        .FileName = Short2Long(FileName)
        .Name = GetParam(sStr, 1)
        .SEOTweak = -Val(GetParam(sStr, 2))
        '.AbsPath = FixFullPath(GetParam(sStr, 2))
        .UserConfigs(0).CompiledPath = GetParam(sStr, 3)
        .UserConfigs(0).HotSpotEditor.HotSpotsFile = GetParam(sStr, 4)
        .FX = Val(GetParam(sStr, 5))
        .NodeExpStatus = GetParam(sStr, 6)
        .ToolBar.BorderColor = Val(GetParam(sStr, 7))
        .CodeOptimization = Val(GetParam(sStr, 8))
        '.UserConfigs(0).HotSpotEditor.IndexFile = GetParam(sStr, 9)
        .version = GetParam(sStr, 10)
        .UserConfigs(0).Frames.UseFrames = -Val(GetParam(sStr, 11))
        .AddIn.Name = GetParam(sStr, 12)
        .AddIn.Description = GetAddInDescription(.AddIn.Name)
        .UserConfigs(0).Frames.FramesFile = GetParam(sStr, 13)
        .UserConfigs(0).RootWeb = GetParam(sStr, 14)
        .SelChangeDelay = Val(GetParam(sStr, 15))
        .ExportHTMLParams = GenExpHTMLPref(GetParam(sStr, 16), .Name, FileName)
        .UserConfigs(0).OptmizePaths = False
        .ToolBar.CreateToolbar = -Val(GetParam(sStr, 16))
        .ToolBar.FollowHScroll = -Val(GetParam(sStr, 17))
        .ToolBar.FollowVScroll = -Val(GetParam(sStr, 18))
        .ToolBar.Alignment = Val(GetParam(sStr, 19))
        .ToolBar.OffsetH = Val(GetParam(sStr, 20))
        .ToolBar.OffsetV = Val(GetParam(sStr, 21))
        .ToolBar.Style = Val(GetParam(sStr, 22))
        .ToolBar.Spanning = Val(GetParam(sStr, 23))
        .ToolBar.bOrder = Val(GetParam(sStr, 24))
        .ToolBar.JustifyHotSpots = -Val(GetParam(sStr, 25))
        .ToolBar.BackColor = Val(GetParam(sStr, 26))
        .ToolBar.Image = GetParam(sStr, 27)
        .ToolBar.CustX = Val(GetParam(sStr, 28))
        .ToolBar.CustY = Val(GetParam(sStr, 29))
        .ToolBar.AttachTo = vbNullString
        .UserConfigs(0).FTP = GetParam(sStr, 30)
        '.FTP.FTPAddress = GetParam(sStr, 30)
        '.FTP.UserName = GetParam(sStr, 31)
        '.FTP.Password = GetParam(sStr, 32)
        '.FTP.ProxyAddress = GetParam(sStr, 33)
        '.FTP.ProxyPort = Val(GetParam(sStr, 34))
        '.FTP.RemoteInfo4FTP = GetParam(sStr, 35)
        .UseGZIP = -Val(GetParam(sStr, 32))
        .BlinkSpeed = Val(GetParam(sStr, 33))
        .BlinkEffect = Val(GetParam(sStr, 34))
        .CustomOffsets = GetParam(sStr, 35)
        .UserConfigs(0).ImagesPath = GetParam(sStr, 36)
        .JSFileName = GetParam(sStr, 37): If LenB(.JSFileName) = 0 Then .JSFileName = "menu"
        ReDim Preserve .UserConfigs(Val(GetParam(sStr, 38)))
        .DefaultConfig = Val(GetParam(sStr, 39))
        
        j = 40: tcp = 14
        For i = 1 To UBound(.UserConfigs)
            With .UserConfigs(i)
                .Name = GetParam(sStr, j + (i - 1) * tcp)
                .Description = GetParam(sStr, j + (i - 1) * tcp + 1)
                .CompiledPath = GetParam(sStr, j + (i - 1) * tcp + 2)
                .Type = Val(GetParam(sStr, j + (i - 1) * tcp + 3))
                .RootWeb = GetParam(sStr, j + (i - 1) * tcp + 4)
                .ImagesPath = GetParam(sStr, j + (i - 1) * tcp + 5)
                .OptmizePaths = -Val(GetParam(sStr, j + (i - 1) * tcp + 6))
                '.Frames.CodeFrame = GetParam(sStr, j + (i - 1) * tcp + 7)
                .FTP = GetParam(sStr, j + (i - 1) * tcp + 7)
                .Frames.FramesFile = GetParam(sStr, j + (i - 1) * tcp + 8)
                '.Frames.MenuFrame = GetParam(sStr, j + (i - 1) * tcp + 9)
                .Frames.UseFrames = -Val(GetParam(sStr, j + (i - 1) * tcp + 10))
                .HotSpotEditor.HotSpotsFile = GetParam(sStr, j + (i - 1) * tcp + 11)
                '.HotSpotEditor.IndexFile = GetParam(sStr, J + (i - 1) * tcp + 12)
                .LocalInfo4RemoteConfig = GetParam(sStr, j + (i - 1) * tcp + 13)
            End With
        Next i
        .UnfoldingSound.onmouseover = GetParam(sStr, j + (i - 1) * tcp)
        .MenusOffset.RootMenusX = Val(GetParam(sStr, j + (i - 1) * tcp + 1))
        .MenusOffset.RootMenusY = Val(GetParam(sStr, j + (i - 1) * tcp + 2))
        .MenusOffset.SubMenusX = Val(GetParam(sStr, j + (i - 1) * tcp + 3))
        .MenusOffset.SubMenusY = Val(GetParam(sStr, j + (i - 1) * tcp + 4))
        
        .LotusDominoSupport = -Val(GetParam(sStr, j + (i - 1) * tcp + 5))
        
        .GenDynAPI = -Val(GetParam(sStr, j + (i - 1) * tcp + 6))
        
        .CompileIECode = -Val(GetParam(sStr, j + (i - 1) * tcp + 7))
        .CompileNSCode = -Val(GetParam(sStr, j + (i - 1) * tcp + 8))
        .CompilehRefFile = -Val(GetParam(sStr, j + (i - 1) * tcp + 9))
        
        .SecondaryProjects = Split(GetParam(sStr, j + (i - 1) * tcp + 10), "|")
        
        .FontSubstitutions = GetParam(sStr, j + (i - 1) * tcp + 11)
        .DoFormsTweak = -Val(GetParam(sStr, j + (i - 1) * tcp + 12))
        
        '.ToolBar.Width = Val(GetParam(sStr, j + (i - 1) * tcp + 13))
        '.ToolBar.Height = Val(GetParam(sStr, j + (i - 1) * tcp + 14))
        
        .StatusTextDisplay = Val(GetParam(sStr, j + (i - 1) * tcp + 13))
        .KeyboardSupport = -Val(GetParam(sStr, j + (i - 1) * tcp + 14))
        .RemoveImageAutoPosCode = -Val(GetParam(sStr, j + (i - 1) * tcp + 15))
        .RootMenusDelay = Val(GetParam(sStr, j + (i - 1) * tcp + 16))
        
        .AnimSpeed = Val(GetParam(sStr, j + (i - 1) * tcp + 17))
        .HideDelay = Val(GetParam(sStr, j + (i - 1) * tcp + 18))
        .SubMenusDelay = Val(GetParam(sStr, j + (i - 1) * tcp + 19))
        
        .DWSupport = -Val(GetParam(sStr, j + (i - 1) * tcp + 20))
        .NS4ClipBug = -Val(GetParam(sStr, j + (i - 1) * tcp + 21))
        
        .UserConfigs(0).Type = Val(GetParam(sStr, j + (i - 1) * tcp + 22))
        
        .AutoSelFunction = -Val(GetParam(sStr, j + (i - 1) * tcp + 23))
        .ImageReadySupport = -Val(GetParam(sStr, j + (i - 1) * tcp + 24))
        
        j = j + (i - 1) * tcp + 25
        If Val(GetParam(sStr, j)) = 0 Then
            .DXFilter = GetParam(sStr, j)
            
            With .AutoScroll
             .maxHeight = GetParam(sStr, j + 1)
             .nColor = GetParam(sStr, j + 2)
             .hColor = GetParam(sStr, j + 3)
             .DnImage.NormalImage = GetParam(sStr, j + 4)
             .DnImage.HoverImage = GetParam(sStr, j + 5)
             .DnImage.w = GetParam(sStr, j + 6)
             .DnImage.h = GetParam(sStr, j + 7)
             .UpImage.NormalImage = GetParam(sStr, j + 8)
             .UpImage.HoverImage = GetParam(sStr, j + 9)
             .margin = GetParam(sStr, j + 10)
             .onmouseover = GetParam(sStr, j + 11)
             .FXhColor = GetParam(sStr, j + 12)
             .FXnColor = GetParam(sStr, j + 13)
             .FXNormal = GetParam(sStr, j + 14)
             .FXOver = GetParam(sStr, j + 15)
             .FXSize = GetParam(sStr, j + 16)
            End With
            
            j = j + 51
        Else
            IsOld = True
        End If
        
        IsPre41 = (.version < 401000)
        IsPre42 = (.version < 402000)
        IsPre45005 = (.version < 405005)
        IsPre49011 = (.version < 409011)
        IsPre49018 = (.version < 409018)
        IsPre42030 = (.version < 420030)
        
        ReDim .Toolbars(Val(GetParam(sStr, j)))
        j = j + 1: tcp = IIf(IsOld, 22, IIf(IsPre41, 23, IIf(IsPre42, 24, IIf(IsPre45005, 27, IIf(IsPre49011, 28, IIf(IsPre49018, 29, IIf(IsPre42030, 30, 34)))))))
        
        For i = 1 To UBound(.Toolbars)
            If parseToolbars Then
                DoEvents
                With .Toolbars(i)
                    .Alignment = Val(GetParam(sStr, j + (i - 1) * tcp))
                    .BackColor = Val(GetParam(sStr, j + (i - 1) * tcp + 1))
                    .bOrder = Val(GetParam(sStr, j + (i - 1) * tcp + 2))
                    .BorderColor = Val(GetParam(sStr, j + (i - 1) * tcp + 3))
                    .ContentsMarginH = Val(GetParam(sStr, j + (i - 1) * tcp + 4))
                    .ContentsMarginV = Val(GetParam(sStr, j + (i - 1) * tcp + 5))
                    .CustX = Val(GetParam(sStr, j + (i - 1) * tcp + 6))
                    .CustY = Val(GetParam(sStr, j + (i - 1) * tcp + 7))
                    .FollowHScroll = -Val(GetParam(sStr, j + (i - 1) * tcp + 8))
                    Select Case Val(GetParam(sStr, j + (i - 1) * tcp + 9))
                        Case 0
                            .FollowVScroll = False
                            .SmartScrolling = False
                        Case 1
                            .FollowVScroll = True
                            .SmartScrolling = False
                        Case 2
                            .FollowVScroll = True
                            .SmartScrolling = True
                        Case 3
                            .FollowVScroll = False
                            .SmartScrolling = True
                    End Select
                    
                    .Groups = Split(GetParam(sStr, j + (i - 1) * tcp + 10), "|")
                    If UBound(.Groups) = -1 Then ReDim .Groups(0)
                    .Height = Val(GetParam(sStr, j + (i - 1) * tcp + 11))
                    .Image = GetParam(sStr, j + (i - 1) * tcp + 12)
                    .JustifyHotSpots = -Val(GetParam(sStr, j + (i - 1) * tcp + 13))
                    .OffsetH = Val(GetParam(sStr, j + (i - 1) * tcp + 14))
                    .OffsetV = Val(GetParam(sStr, j + (i - 1) * tcp + 15))
                    .Name = GetParam(sStr, j + (i - 1) * tcp + 16)
                    .Separation = Val(GetParam(sStr, j + (i - 1) * tcp + 17))
                    .Spanning = Val(GetParam(sStr, j + (i - 1) * tcp + 18))
                    .Style = Val(GetParam(sStr, j + (i - 1) * tcp + 19))
                    .Width = Val(GetParam(sStr, j + (i - 1) * tcp + 20))
                    If InStr(GetParam(sStr, j + (i - 1) * tcp + 21), "|") Then
                        .AttachTo = Split(GetParam(sStr, j + (i - 1) * tcp + 21), "|")(0)
                        .AttachToAlignment = Val(Split(GetParam(sStr, j + (i - 1) * tcp + 21), "|")(1))
                    Else
                        .AttachTo = GetParam(sStr, j + (i - 1) * tcp + 21)
                    End If
                    If Not IsOld Then
                        .Condition = GetParam(sStr, j + (i - 1) * tcp + 22)
                        If Not IsPre41 Then .BorderStyle = Val(GetParam(sStr, j + (i - 1) * tcp + 23))
                        If Not IsPre42 Then
                            .DropShadowColor = Val(GetParam(sStr, j + (i - 1) * tcp + 24))
                            .DropShadowSize = Val(GetParam(sStr, j + (i - 1) * tcp + 25))
                            .Transparency = Val(GetParam(sStr, j + (i - 1) * tcp + 26))
                        End If
                        If Not IsPre45005 Then
                            #If DEVVER = 1 Then
                            .IsTemplate = -Val(GetParam(sStr, j + (i - 1) * tcp + 27))
                            If .IsTemplate Then
                                If .Name <> "DynAPI_TBTemplate" Then .IsTemplate = False
                            End If
                            #Else
                            .IsTemplate = 0
                            #End If
                            
                            .AttachToAutoResize = -Val(GetParam(sStr, j + (i - 1) * tcp + 28))
                            
                            If Not IsPre49018 Then
                                .Compile = -Val(GetParam(sStr, j + (i - 1) * tcp + 29))
                            End If
                            
                            If Not IsPre42030 Then
                                .Radius.TopLeft = Val(GetParam(sStr, j + (i - 1) * tcp + 30))
                                .Radius.TopRight = Val(GetParam(sStr, j + (i - 1) * tcp + 31))
                                .Radius.BottomLeft = Val(GetParam(sStr, j + (i - 1) * tcp + 32))
                                .Radius.BottomRight = Val(GetParam(sStr, j + (i - 1) * tcp + 33))
                            End If
                        End If
                    End If
                End With
            End If
        Next i
        
        If IsOld Then .DXFilter = GetParam(sStr, j + (i - 1) * tcp + 24)
        
        LoadAddInParams .AddIn.Name
        
        If .DefaultConfig > UBound(.UserConfigs) Then .DefaultConfig = UBound(.UserConfigs)
    End With
    
End Function

Public Sub LoadAddInParams(AddIn As String)

    Dim ff As Integer
    Dim i As Integer
    Dim FileName As String
    Dim tmpStr As String
    
    On Error GoTo ExitSub
    
    Erase params
    
    FileName = AppPath + "AddIns\" + AddIn + ".par"
    If AddIn <> "" And FileExists(FileName) Then
        ff = FreeFile
        Open FileName For Input As #ff
            i = 1
            Do Until EOF(ff)
                ReDim Preserve params(i)
                With params(i)
                    Line Input #ff, params(i).Name
                    Line Input #ff, params(i).Description
                    Line Input #ff, params(i).Default
                    Line Input #ff, params(i).Value
                    Line Input #ff, tmpStr
                    params(i).Required = -CInt(tmpStr)
                End With
                i = i + 1
            Loop
        Close #ff: ff = 0
    Else
        ReDim params(0)
    End If
    
    Exit Sub
    
ExitSub:
    If ff > 0 Then Close #ff
    MsgBox "Error loading the parameters for the " + AddIn + " AddIn" + vbCrLf + vbCrLf + _
            "Error " & Err.number & ": " + Err.Description, vbCritical + vbOKOnly, GetLocalizedStr(665)
    
End Sub

Public Sub SaveAddInParams(AddIn As String)

    Dim ff As Integer
    Dim i As Integer
    Dim FileName As String
    
    On Error GoTo ExitSub

    FileName = AppPath + "AddIns\" + AddIn + ".par"
    If LenB(Dir(FileName)) <> 0 Then Kill FileName
    If UBound(params) > 0 Then
        ff = FreeFile
        Open FileName For Output As #ff
            For i = 1 To UBound(params)
                Print #ff, params(i).Name
                Print #ff, params(i).Description
                Print #ff, params(i).Default
                Print #ff, params(i).Value
                Print #ff, Abs(params(i).Required)
            Next i
        Close #ff: ff = 0
    End If
    
Exit Sub
    
ExitSub:
    If ff > 0 Then Close #ff
    MsgBox "Error saving the parameters for the " + AddIn + " AddIn" + vbCrLf + vbCrLf + _
            "Error " & Err.number & ": " + Err.Description, vbCritical + vbOKOnly, GetLocalizedStr(666)
    
End Sub

Public Sub AddCommandSeparator(Optional params As String)

    Dim NewItm As Node
    Dim prnt As Node
    Dim i As Integer
    Dim NewCmd As MenuCmd
    
    On Error GoTo ExitSub
    
    Set prnt = GetRealParent
    
    ReDim Preserve MenuCmds(UBound(MenuCmds) + 1)
    With MenuCmds(UBound(MenuCmds))
        .Name = "[SEP]"
        If LenB(params) = 0 Then
            .nTextColor = &H0
            .nBackColor = -2
            .SeparatorPercent = 80
            .Compile = True
        Else
            .nTextColor = Val(GetParam(params, 6))
            .nBackColor = Val(GetParam(params, 5))
            .SeparatorPercent = Val(GetParam(params, 58))
            .Compile = Val(GetParam(params, 66))
        End If
        .parent = GetID(prnt)
    End With
    
    If Left$(frmMain.tvMenus.SelectedItem.key, 1) <> "G" Then
        i = UBound(MenuCmds)
        NewCmd = MenuCmds(i)
        Do Until (Left$(frmMain.tvMenus.SelectedItem.key, 1) & (i - 1)) = frmMain.tvMenus.SelectedItem.key
            MenuCmds(i) = MenuCmds(i - 1)
            If MenuCmds(i).Name = "[SEP]" Then
                frmMain.tvMenus.Nodes("S" & i - 1).key = "S" & i
            Else
                frmMain.tvMenus.Nodes("C" & i - 1).key = "C" & i
            End If
            i = i - 1
        Loop
        MenuCmds(i) = NewCmd
        Set NewItm = frmMain.tvMenus.Nodes.Add(frmMain.tvMenus.SelectedItem.key, tvwNext, "S" & i)
    Else
        Set NewItm = frmMain.tvMenus.Nodes.Add("G" & GetID(prnt), tvwChild, "S" & UBound(MenuCmds))
    End If
    With NewItm
        .Text = String$(10, "-")
        .Image = IconIndex("Separator")
        .Selected = True
        .EnsureVisible
        If MenuGrps(MenuCmds(UBound(MenuCmds)).parent).IsTemplate Then .ForeColor = vbBlue
    End With
    
    #If ISCOMP = 0 Then
    frmMain.UpdateControls
    #End If
    
ExitSub:

End Sub

Public Function FixItemName(ByVal ItemName As String) As String

    Dim i As Integer
    Dim m As String
    
    ItemName = Replace(ItemName, " ", "_")
    ItemName = Replace(ItemName, "-", "_")
    
    ItemName = Replace(ItemName, "á", "a")
    ItemName = Replace(ItemName, "é", "e")
    ItemName = Replace(ItemName, "í", "i")
    ItemName = Replace(ItemName, "ó", "o")
    ItemName = Replace(ItemName, "ú", "u")
     
    ItemName = Replace(ItemName, "Á", "A")
    ItemName = Replace(ItemName, "É", "E")
    ItemName = Replace(ItemName, "Í", "I")
    ItemName = Replace(ItemName, "Ó", "O")
    ItemName = Replace(ItemName, "Ú", "U")
    
    ItemName = Replace(ItemName, "à", "a")
    ItemName = Replace(ItemName, "è", "e")
    ItemName = Replace(ItemName, "ì", "i")
    ItemName = Replace(ItemName, "ò", "o")
    ItemName = Replace(ItemName, "ù", "u")
    
    ItemName = Replace(ItemName, "À", "A")
    ItemName = Replace(ItemName, "È", "E")
    ItemName = Replace(ItemName, "Ì", "I")
    ItemName = Replace(ItemName, "Ò", "O")
    ItemName = Replace(ItemName, "Ù", "U")
    
    ItemName = Replace(ItemName, "ä", "a")
    ItemName = Replace(ItemName, "ë", "e")
    ItemName = Replace(ItemName, "ï", "i")
    ItemName = Replace(ItemName, "ö", "o")
    ItemName = Replace(ItemName, "ü", "u")
    
    ItemName = Replace(ItemName, "Ä", "A")
    ItemName = Replace(ItemName, "Ë", "E")
    ItemName = Replace(ItemName, "Ï", "I")
    ItemName = Replace(ItemName, "Ö", "O")
    ItemName = Replace(ItemName, "Ü", "U")
    
    For i = 1 To Len(ItemName)
        m = LCase(Mid(ItemName, i, 1))
        If Not (m >= "0" And m <= "9" Or _
            m >= "a" And m <= "z" Or _
            m = "-") Then
            Mid(ItemName, i, 1) = "_"
        End If
    Next i
    
    FixItemName = ItemName
    
End Function

Private Function GetRealParent() As Node
    
    Dim prnt As Node
    Dim tMenu As Integer

    Set prnt = frmMain.tvMenus.SelectedItem
    If Not prnt.parent Is Nothing Then
        Set prnt = prnt.parent
    End If
    With MenuGrps(GetID(prnt)).Actions
        If .onmouseover.Type = atcCascade Then
            tMenu = .onmouseover.TargetMenu
        Else
            If .onclick.Type = atcCascade Then
                tMenu = .onclick.TargetMenu
            Else
                If .OnDoubleClick.Type = atcCascade Then
                    tMenu = .OnDoubleClick.TargetMenu
                End If
            End If
        End If
    End With
    If tMenu <> 0 Then Set prnt = frmMain.tvMenus.Nodes("G" & tMenu)
    
    Set GetRealParent = prnt

End Function

Public Sub AddMenuCommand(Optional params As String, Optional CancelUpdate As Boolean, Optional CancelRename As Boolean, Optional UseParentFromParams As Boolean)

    Dim NewItm As Node
    Dim prnt As Node
    Dim i As Integer
    Dim NewCmd As MenuCmd
    Dim ThisConfig As ConfigDef
    Dim FirstCmdIsSep As Boolean
    Dim tmpName As String
    Dim rID As Integer
    
    On Error GoTo ExitSub
    
    If LenB(params) <> 0 Then
        If GetParam(params, 1) = "[SEP]" Then
            AddCommandSeparator params
            Exit Sub
        End If
    End If
    
    If UseParentFromParams Then
        Set prnt = frmMain.tvMenus.Nodes("G" + GetParam(params, 19))
    Else
        Set prnt = GetRealParent
    End If
    
    With MenuGrps(GetID(prnt)).Actions
        If .onclick.Type <> atcCascade And .OnDoubleClick.Type <> atcCascade And .onmouseover.Type <> atcCascade Then
            .onmouseover.Type = atcCascade
            .onmouseover.TargetMenu = GetID(prnt)
        End If
    End With
    
    ReDim Preserve MenuCmds(UBound(MenuCmds) + 1)
    With MenuCmds(UBound(MenuCmds))
        If LenB(params) = 0 Then
            FirstCmdIsSep = False
            If prnt.children > 0 Then FirstCmdIsSep = IsSeparator(prnt.Child.FirstSibling.key)
            If Preferences.CommandsInheritance = icDefault Or prnt.children = 0 Or FirstCmdIsSep Then
                MenuCmds(UBound(MenuCmds)) = TemplateCommand
                rID = SubMenuOf(GetID(prnt))
                If rID > 0 Then
                    .nTextColor = MenuCmds(rID).nTextColor
                    .nBackColor = MenuCmds(rID).nBackColor
                    .hTextColor = MenuCmds(rID).hTextColor
                    .hBackColor = MenuCmds(rID).hBackColor
                    .NormalFont = MenuCmds(rID).NormalFont
                    .HoverFont = MenuCmds(rID).HoverFont
                    .BackImage = MenuCmds(rID).BackImage
                    .iCursor = MenuCmds(rID).iCursor
                Else
                    rID = GetID(prnt)
                    .nTextColor = MenuGrps(rID).nTextColor
                    If CreateToolbar Then
                        .nBackColor = MenuGrps(rID).nBackColor
                    Else
                        .nBackColor = MenuGrps(rID).bColor
                    End If
                    .hTextColor = MenuGrps(rID).hTextColor
                    .hBackColor = MenuGrps(rID).hBackColor
                    .NormalFont = MenuGrps(rID).DefNormalFont
                    .HoverFont = MenuGrps(rID).DefHoverFont
                End If
                .parent = GetID(prnt)
                ThisConfig = Project.UserConfigs(Project.DefaultConfig)
            Else
                MenuCmds(UBound(MenuCmds)) = MenuCmds(GetID(prnt.Child.FirstSibling))
                MenuCmds(UBound(MenuCmds)).Actions = TemplateCommand.Actions
                .caption = vbNullString
                .WinStatus = vbNullString
            End If
            .Name = vbNullString
            .disabled = False
            .Actions.onmouseover.TargetMenuAlignment = IIf(MenuGrps(.parent).AlignmentStyle = ascHorizontal, gacBottomLeft, gacRightTop)
            .Actions.onclick.TargetMenuAlignment = IIf(MenuGrps(.parent).AlignmentStyle = ascHorizontal, gacBottomLeft, gacRightTop)
            .Actions.OnDoubleClick.TargetMenuAlignment = IIf(MenuGrps(.parent).AlignmentStyle = ascHorizontal, gacBottomLeft, gacRightTop)
        Else
            tmpName = FixItemName(GetParam(params, 1))
            If LenB(tmpName) = 0 Or ItemExists(tmpName, False) Then tmpName = GetSecuenceName(False, "Command")
            .Name = tmpName
            
            .caption = GetParam(params, 2)
            .WinStatus = GetParam(params, 21)
            .iCursor.cType = GetParam(params, 18)
            .iCursor.CFile = GetParam(params, 20)
            .parent = GetID(prnt)
            
            'FOR COMPATIBILITY
            If LenB(GetParam(params, 17)) <> 0 Or Val(GetParam(params, 23)) > 0 Then 'Is old Format
                Select Case Val(GetParam(params, 20))
                    Case ByClicking
                        If -Val(GetParam(params, 22)) Then
                            .Actions.onclick.TargetMenu = Val(GetParam(params, 23))
                            .Actions.onclick.Type = atcCascade
                        Else
                            .Actions.onclick.url = GetParam(params, 17)
                            .Actions.onclick.Type = atcURL
                        End If
                    Case ByHovering
                        If -Val(GetParam(params, 22)) Then
                            .Actions.onmouseover.TargetMenu = Val(GetParam(params, 23))
                            .Actions.onmouseover.Type = atcCascade
                        Else
                            .Actions.onmouseover.url = GetParam(params, 17)
                            .Actions.onmouseover.Type = atcURL
                        End If
                End Select
            End If
            
            .Alignment = Val(GetParam(params, 29))
            '.TargetFrame = GetParam(Params, 31)
            .disabled = -Val(GetParam(params, 34))
            .hBackColor = Val(GetParam(params, 3))
            .hTextColor = Val(GetParam(params, 4))
            .nBackColor = Val(GetParam(params, 5))
            .nTextColor = Val(GetParam(params, 6))
            
            .HoverFont.FontName = GetParam(params, 7)
            .HoverFont.FontSize = Val(GetParam(params, 8))
            .HoverFont.FontBold = -Val(GetParam(params, 9))
            .HoverFont.FontItalic = -Val(GetParam(params, 10))
            .HoverFont.FontUnderline = -Val(GetParam(params, 11))
            
            .NormalFont.FontName = GetParam(params, 12)
            .NormalFont.FontSize = Val(GetParam(params, 13))
            .NormalFont.FontBold = -Val(GetParam(params, 14))
            .NormalFont.FontItalic = -Val(GetParam(params, 15))
            .NormalFont.FontUnderline = -Val(GetParam(params, 16))
            
            .LeftImage.NormalImage = GetParam(params, 24)
            .LeftImage.HoverImage = GetParam(params, 25)
            .LeftImage.w = Val(GetParam(params, 27))
            .LeftImage.h = Val(GetParam(params, 28))
            .LeftImage.margin = Val(GetParam(params, 91))
            
            .RightImage.NormalImage = GetParam(params, 35)
            .RightImage.HoverImage = GetParam(params, 36)
            .RightImage.w = Val(GetParam(params, 37))
            .RightImage.h = Val(GetParam(params, 38))
            .RightImage.margin = Val(GetParam(params, 92))
            
            .BackImage.NormalImage = GetParam(params, 51)
            .BackImage.HoverImage = GetParam(params, 52)
            .BackImage.Tile = Val(GetParam(params, 67))
            .BackImage.AllowCrop = Val(GetParam(params, 68))
            .BackImage.w = Val(GetParam(params, 69))
            .BackImage.h = Val(GetParam(params, 70))
            
            If Val(GetParam(params, 26)) = 1 Then
                'Upgrade from old versions
                .RightImage.NormalImage = .LeftImage.NormalImage
                .RightImage.HoverImage = .LeftImage.HoverImage
                .RightImage.w = .LeftImage.w
                .RightImage.h = .LeftImage.h
                .LeftImage.NormalImage = vbNullString
                .LeftImage.HoverImage = vbNullString
                .LeftImage.w = 0
                .LeftImage.h = 0
            End If
            If LenB(.RightImage.NormalImage) = 0 Then .RightImage.w = 0: .RightImage.h = 0
            If LenB(.LeftImage.NormalImage) = 0 Then .LeftImage.w = 0: .LeftImage.h = 0
            
            'Check if this its an old project
            'to avoid overwriting stuff...
            If LenB(GetParam(params, 17)) = 0 And Val(GetParam(params, 23)) = 0 Then 'Is not old Format
            With .Actions.onclick
                .Type = Val(GetParam(params, 39))
                .url = GetParam(params, 40)
                .TargetFrame = GetParam(params, 41)
                .TargetMenu = Val(GetParam(params, 42))
                .WindowOpenParams = Replace(GetParam(params, 53), "|", cSep)
                .TargetMenuAlignment = Val(GetParam(params, 31))
            End With
            With .Actions.onmouseover
                .Type = Val(GetParam(params, 43))
                .url = GetParam(params, 44)
                .TargetFrame = GetParam(params, 45)
                .TargetMenu = Val(GetParam(params, 46))
                .WindowOpenParams = Replace(GetParam(params, 54), "|", cSep)
                .TargetMenuAlignment = Val(GetParam(params, 32))
            End With
            With .Actions.OnDoubleClick
                .Type = Val(GetParam(params, 47))
                .url = GetParam(params, 48)
                .TargetFrame = GetParam(params, 49)
                .TargetMenu = Val(GetParam(params, 50))
                .WindowOpenParams = Replace(GetParam(params, 55), "|", cSep)
                .TargetMenuAlignment = Val(GetParam(params, 33))
            End With
            End If
            
            .Sound.onmouseover = GetParam(params, 56)
            .Sound.onclick = GetParam(params, 57)
            
            .CmdsFXhColor = Val(GetParam(params, 59))
            .CmdsFXnColor = Val(GetParam(params, 60))
            .CmdsFXNormal = Val(GetParam(params, 61))
            .CmdsFXOver = Val(GetParam(params, 62))
            .CmdsFXSize = Val(GetParam(params, 63))
            .CmdsMarginX = Val(GetParam(params, 64))
            .CmdsMarginY = Val(GetParam(params, 65))
            
            .Compile = Val(GetParam(params, 66))
            
            .NormalFont.FontShadow.Enabled1 = Val(GetParam(params, 71))
            .NormalFont.FontShadow.Color1 = Val(GetParam(params, 72))
            .NormalFont.FontShadow.OffsetX1 = Val(GetParam(params, 73))
            .NormalFont.FontShadow.OffsetY1 = Val(GetParam(params, 74))
            .NormalFont.FontShadow.Blur1 = Val(GetParam(params, 75))
            .NormalFont.FontShadow.Enabled2 = Val(GetParam(params, 76))
            .NormalFont.FontShadow.Color2 = Val(GetParam(params, 77))
            .NormalFont.FontShadow.OffsetX2 = Val(GetParam(params, 78))
            .NormalFont.FontShadow.OffsetY2 = Val(GetParam(params, 79))
            .NormalFont.FontShadow.Blur2 = Val(GetParam(params, 80))
            
            .HoverFont.FontShadow.Enabled1 = Val(GetParam(params, 81))
            .HoverFont.FontShadow.Color1 = Val(GetParam(params, 82))
            .HoverFont.FontShadow.OffsetX1 = Val(GetParam(params, 83))
            .HoverFont.FontShadow.OffsetY1 = Val(GetParam(params, 84))
            .HoverFont.FontShadow.Blur1 = Val(GetParam(params, 85))
            .HoverFont.FontShadow.Enabled2 = Val(GetParam(params, 86))
            .HoverFont.FontShadow.Color2 = Val(GetParam(params, 87))
            .HoverFont.FontShadow.OffsetX2 = Val(GetParam(params, 88))
            .HoverFont.FontShadow.OffsetY2 = Val(GetParam(params, 89))
            .HoverFont.FontShadow.Blur2 = Val(GetParam(params, 90))
            
            .Radius.TopLeft = Val(GetParam(params, 96))
            .Radius.TopRight = Val(GetParam(params, 94))
            .Radius.BottomLeft = Val(GetParam(params, 95))
            .Radius.BottomRight = Val(GetParam(params, 96))
        End If
    End With
    
    If Left$(frmMain.tvMenus.SelectedItem.key, 1) <> "G" Then
        i = UBound(MenuCmds)
        NewCmd = MenuCmds(i)
        Do Until (Left$(frmMain.tvMenus.SelectedItem.key, 1) & (i - 1)) = frmMain.tvMenus.SelectedItem.key
            MenuCmds(i) = MenuCmds(i - 1)
            If MenuCmds(i).Name = "[SEP]" Then
                frmMain.tvMenus.Nodes("S" & i - 1).key = "S" & i
            Else
                frmMain.tvMenus.Nodes("C" & i - 1).key = "C" & i
            End If
            i = i - 1
        Loop
        MenuCmds(i) = NewCmd
        Set NewItm = frmMain.tvMenus.Nodes.Add(frmMain.tvMenus.SelectedItem.key, tvwNext, "C" & i)
    Else
        Set NewItm = frmMain.tvMenus.Nodes.Add("G" & GetID(prnt), tvwChild, "C" & UBound(MenuCmds))
    End If
    NewItm.Image = GenCmdIcon(UBound(MenuCmds))
    With NewItm
        .Selected = True
        If Not CancelUpdate Then .EnsureVisible
        If MenuGrps(MenuCmds(UBound(MenuCmds)).parent).IsTemplate Then .ForeColor = vbBlue
    End With
    If LenB(params) = 0 Then
        If Not CancelRename Then frmMain.tvMenus.StartLabelEdit
    Else
        NewItm.Text = MenuCmds(GetID).Name
    End If
    
    #If ISCOMP = 0 Then
    If Not CancelUpdate Then frmMain.UpdateControls
    #End If
    
ExitSub:

End Sub

Public Sub AddMenuGroup(Optional params As String, Optional CancelRename As Boolean)

    Dim NewItm As Node
    Dim AltTemplate As MenuGrp
    Dim ThisConfig As ConfigDef
    Dim tmpName As String

    ReDim Preserve MenuGrps(UBound(MenuGrps) + 1)
    With MenuGrps(UBound(MenuGrps))
        If LenB(params) = 0 Then
            If Preferences.GroupsInheritance = icDefault Or frmMain.tvMenus.Nodes.Count = 0 Then
                AltTemplate = TemplateGroup
                ThisConfig = Project.UserConfigs(Project.DefaultConfig)
            Else
                #If DEVVER = 1 Then
                    If MenuGrps(GetID(frmMain.tvMenus.Nodes(1))).IsTemplate Then
                        AltTemplate = TemplateGroup
                        ThisConfig = Project.UserConfigs(Project.DefaultConfig)
                    Else
                        AltTemplate = MenuGrps(GetID(frmMain.tvMenus.Nodes(1)))
                    End If
                #Else
                    AltTemplate = MenuGrps(GetID(frmMain.tvMenus.Nodes(1)))
                #End If
            End If
            MenuGrps(UBound(MenuGrps)) = AltTemplate
            .Name = vbNullString
            .caption = vbNullString
            .Actions = TemplateGroup.Actions
            .disabled = False
        Else
            On Error Resume Next
            tmpName = FixItemName(GetParam(params, 1))
            If LenB(tmpName) = 0 Or ItemExists(tmpName, True) Then tmpName = GetSecuenceName(True, "Group")
            .Name = tmpName
            
            .Leading = Val(GetParam(params, 19))
            .bColor = GetParam(params, 2)
            .Corners.rightCorner = GetParam(params, 3)
            .Corners.leftCorner = GetParam(params, 4)
            .Corners.bottomCorner = GetParam(params, 75)
            .Corners.topCorner = GetParam(params, 76)
            .frameBorder = GetParam(params, 6)
            
            .DefHoverFont.FontName = GetParam(params, 7)
            .DefHoverFont.FontSize = GetParam(params, 8)
            .DefHoverFont.FontBold = GetParam(params, 9)
            .DefHoverFont.FontItalic = GetParam(params, 10)
            .DefHoverFont.FontUnderline = GetParam(params, 11)
            .DefNormalFont.FontName = GetParam(params, 12)
            .DefNormalFont.FontSize = GetParam(params, 13)
            .DefNormalFont.FontBold = GetParam(params, 14)
            .DefNormalFont.FontItalic = GetParam(params, 15)
            .DefNormalFont.FontUnderline = GetParam(params, 16)
            
            .x = Val(GetParam(params, 17))
            .y = Val(GetParam(params, 18))
            
            .Image = GetParam(params, 20)
            
            .Alignment = Val(GetParam(params, 21))
            
            .DropShadowColor = Val(GetParam(params, 22))
            .DropShadowSize = Val(GetParam(params, 96))
            .Transparency = Val(GetParam(params, 23))
            
            .CmdsFXNormal = Val(GetParam(params, 24))
            .CmdsFXOver = Val(GetParam(params, 25))
            .CmdsFXSize = Val(GetParam(params, 26))
            .CmdsMarginX = Val(GetParam(params, 59))
            .CmdsMarginY = Val(GetParam(params, 60))
            
            .CmdsFXnColor = Val(GetParam(params, 71))
            .CmdsFXhColor = Val(GetParam(params, 72))
            
            .iCursor.cType = Val(GetParam(params, 27))
            .iCursor.CFile = GetParam(params, 5)
            
            .fWidth = Val(GetParam(params, 28))
            .fHeight = Val(GetParam(params, 73))
            
            .CaptionAlignment = Val(GetParam(params, 29))
            
            .tbiLeftImage.NormalImage = GetParam(params, 30)
            .tbiLeftImage.HoverImage = GetParam(params, 31)
            .tbiLeftImage.w = Val(GetParam(params, 32))
            .tbiLeftImage.h = Val(GetParam(params, 33))
            .tbiLeftImage.margin = Val(GetParam(params, 130))
            
            .tbiRightImage.NormalImage = GetParam(params, 34)
            .tbiRightImage.HoverImage = GetParam(params, 35)
            .tbiRightImage.w = Val(GetParam(params, 36))
            .tbiRightImage.h = Val(GetParam(params, 37))
            .tbiRightImage.margin = Val(GetParam(params, 131))
            
            .tbiBackImage.NormalImage = GetParam(params, 97)
            .tbiBackImage.HoverImage = GetParam(params, 98)
            .tbiBackImage.Tile = Val(GetParam(params, 106))
            .tbiBackImage.AllowCrop = Val(GetParam(params, 107))
            .tbiBackImage.w = Val(GetParam(params, 108))
            .tbiBackImage.h = Val(GetParam(params, 109))
            
            .hBackColor = Val(GetParam(params, 38))
            .hTextColor = Val(GetParam(params, 39))
            .nBackColor = Val(GetParam(params, 40))
            .nTextColor = Val(GetParam(params, 41))
            
            .disabled = -Val(GetParam(params, 42))
            
            .caption = GetParam(params, 43)
            
            .IncludeInToolbar = -Val(GetParam(params, 44))
            .ToolbarIndex = Val(GetParam(params, 45))
            
            .WinStatus = GetParam(params, 46)
            
            .Actions.onclick.Type = Val(GetParam(params, 47))
            .Actions.onclick.url = GetParam(params, 48)
            .Actions.onclick.TargetFrame = GetParam(params, 49)
            .Actions.onclick.TargetMenu = Val(GetParam(params, 50))
            .Actions.onclick.WindowOpenParams = Replace(GetParam(params, 63), "|", cSep)
            
            .Actions.onmouseover.Type = Val(GetParam(params, 51))
            .Actions.onmouseover.url = GetParam(params, 52)
            .Actions.onmouseover.TargetFrame = GetParam(params, 53)
            .Actions.onmouseover.TargetMenu = Val(GetParam(params, 54))
            .Actions.onmouseover.WindowOpenParams = Replace(GetParam(params, 64), "|", cSep)
            
            .Actions.OnDoubleClick.Type = Val(GetParam(params, 55))
            .Actions.OnDoubleClick.url = GetParam(params, 56)
            .Actions.OnDoubleClick.TargetFrame = GetParam(params, 57)
            .Actions.OnDoubleClick.TargetMenu = Val(GetParam(params, 58))
            .Actions.OnDoubleClick.WindowOpenParams = Replace(GetParam(params, 65), "|", cSep)
            
            .ContentsMarginH = Val(GetParam(params, 61))
            .ContentsMarginV = Val(GetParam(params, 62))
            
            .HSImage = GetParam(params, 66)
            
            .IsContext = -Val(GetParam(params, 67))
            
            .Sound.onmouseover = GetParam(params, 68)
            .Sound.onclick = GetParam(params, 69)
            
            .BorderStyle = GetParam(params, 70)
            
            .IsTemplate = -Val(GetParam(params, 74))
            
            .AlignmentStyle = Val(GetParam(params, 77))
            
            .CornersImages.gcTopLeft = GetParam(params, 78)
            .CornersImages.gcTopCenter = GetParam(params, 79)
            .CornersImages.gcTopRight = GetParam(params, 80)
            .CornersImages.gcLeft = GetParam(params, 81)
            .CornersImages.gcRight = GetParam(params, 82)
            .CornersImages.gcBottomLeft = GetParam(params, 83)
            .CornersImages.gcBottomCenter = GetParam(params, 84)
            .CornersImages.gcBottomRight = GetParam(params, 85)
            
            .scrolling.maxHeight = GetParam(params, 86)
            .scrolling.nColor = GetParam(params, 87)
            .scrolling.hColor = GetParam(params, 88)
            .scrolling.DnImage.NormalImage = GetParam(params, 89)
            .scrolling.DnImage.HoverImage = GetParam(params, 90)
            .scrolling.DnImage.w = GetParam(params, 91)
            .scrolling.DnImage.h = GetParam(params, 92)
            .scrolling.UpImage.NormalImage = GetParam(params, 93)
            .scrolling.UpImage.HoverImage = GetParam(params, 94)
            .scrolling.margin = GetParam(params, 95)
            .scrolling.onmouseover = -Val(GetParam(params, 99))
            .scrolling.FXhColor = Val(GetParam(params, 100))
            .scrolling.FXnColor = Val(GetParam(params, 101))
            .scrolling.FXNormal = Val(GetParam(params, 102))
            .scrolling.FXOver = Val(GetParam(params, 103))
            .scrolling.FXSize = Val(GetParam(params, 104))
            
            .Compile = Val(GetParam(params, 105))
            
            .DefNormalFont.FontShadow.Enabled1 = Val(GetParam(params, 110))
            .DefNormalFont.FontShadow.Color1 = Val(GetParam(params, 111))
            .DefNormalFont.FontShadow.OffsetX1 = Val(GetParam(params, 112))
            .DefNormalFont.FontShadow.OffsetY1 = Val(GetParam(params, 113))
            .DefNormalFont.FontShadow.Blur1 = Val(GetParam(params, 114))
            .DefNormalFont.FontShadow.Enabled2 = Val(GetParam(params, 115))
            .DefNormalFont.FontShadow.Color2 = Val(GetParam(params, 116))
            .DefNormalFont.FontShadow.OffsetX2 = Val(GetParam(params, 117))
            .DefNormalFont.FontShadow.OffsetY2 = Val(GetParam(params, 118))
            .DefNormalFont.FontShadow.Blur2 = Val(GetParam(params, 119))
            
            .DefHoverFont.FontShadow.Enabled1 = Val(GetParam(params, 120))
            .DefHoverFont.FontShadow.Color1 = Val(GetParam(params, 121))
            .DefHoverFont.FontShadow.OffsetX1 = Val(GetParam(params, 122))
            .DefHoverFont.FontShadow.OffsetY1 = Val(GetParam(params, 123))
            .DefHoverFont.FontShadow.Blur1 = Val(GetParam(params, 124))
            .DefHoverFont.FontShadow.Enabled2 = Val(GetParam(params, 125))
            .DefHoverFont.FontShadow.Color2 = Val(GetParam(params, 126))
            .DefHoverFont.FontShadow.OffsetX2 = Val(GetParam(params, 127))
            .DefHoverFont.FontShadow.OffsetY2 = Val(GetParam(params, 128))
            .DefHoverFont.FontShadow.Blur2 = Val(GetParam(params, 129))
            
            .Radius.TopLeft = Val(GetParam(params, 132))
            .Radius.TopRight = Val(GetParam(params, 133))
            .Radius.BottomLeft = Val(GetParam(params, 134))
            .Radius.BottomRight = Val(GetParam(params, 135))
            
            .tbiRadius.TopLeft = Val(GetParam(params, 136))
            .tbiRadius.TopRight = Val(GetParam(params, 137))
            .tbiRadius.BottomLeft = Val(GetParam(params, 138))
            .tbiRadius.BottomRight = Val(GetParam(params, 139))
        End If
    End With
    
    Set NewItm = frmMain.tvMenus.Nodes.Add(, , "G" & UBound(MenuGrps))
    NewItm.Image = GenGrpIcon(UBound(MenuGrps))
    With NewItm
        .Selected = True
        If Not CancelRename Then .EnsureVisible
        .Bold = True
        If MenuGrps(UBound(MenuGrps)).IsTemplate Then .ForeColor = vbBlue
    End With
    
    If LenB(params) = 0 Then
        If Not CancelRename Then frmMain.tvMenus.StartLabelEdit
    Else
        NewItm.Text = MenuGrps(GetID).Name
    End If
    
    #If ISCOMP = 0 Then
    frmMain.UpdateControls
    #End If
    
End Sub

Public Function GetGrpParams(grp As MenuGrp) As String

    Dim sStr As String

    With grp
        sStr = .Name                                                '1
        sStr = sStr + cSep & .bColor                                '2
        sStr = sStr + cSep & .Corners.rightCorner                   '3
        sStr = sStr + cSep & .Corners.leftCorner                    '4
        sStr = sStr + cSep & .iCursor.CFile                         '5
        sStr = sStr + cSep & .frameBorder                           '6
        
        sStr = sStr + cSep & .DefHoverFont.FontName                 '7
        sStr = sStr + cSep & .DefHoverFont.FontSize                 '8
        sStr = sStr + cSep & Abs(.DefHoverFont.FontBold)            '9
        sStr = sStr + cSep & Abs(.DefHoverFont.FontItalic)          '10
        sStr = sStr + cSep & Abs(.DefHoverFont.FontUnderline)       '11
        sStr = sStr + cSep & .DefNormalFont.FontName                '12
        sStr = sStr + cSep & .DefNormalFont.FontSize                '13
        sStr = sStr + cSep & Abs(.DefNormalFont.FontBold)           '14
        sStr = sStr + cSep & Abs(.DefNormalFont.FontItalic)         '15
        sStr = sStr + cSep & Abs(.DefNormalFont.FontUnderline)      '16
        
        sStr = sStr + cSep & .x                                     '17
        sStr = sStr + cSep & .y                                     '18
        
        sStr = sStr + cSep & .Leading                               '19
        
        sStr = sStr + cSep + .Image                                 '20
        
        sStr = sStr + cSep & .Alignment                             '21
        
        sStr = sStr + cSep & .DropShadowColor                       '22
        sStr = sStr + cSep & .Transparency                          '23
        
        sStr = sStr + cSep & .CmdsFXNormal                          '24
        sStr = sStr + cSep & .CmdsFXOver                            '25
        sStr = sStr + cSep & .CmdsFXSize                            '26
        
        sStr = sStr + cSep & .iCursor.cType                         '27
        
        sStr = sStr + cSep & .fWidth                                '28
        
        sStr = sStr + cSep & .CaptionAlignment                      '29
        
        sStr = sStr + cSep & .tbiLeftImage.NormalImage              '30
        sStr = sStr + cSep & .tbiLeftImage.HoverImage               '31
        sStr = sStr + cSep & .tbiLeftImage.w                        '32
        sStr = sStr + cSep & .tbiLeftImage.h                        '33
        sStr = sStr + cSep & .tbiRightImage.NormalImage             '34
        sStr = sStr + cSep & .tbiRightImage.HoverImage              '35
        sStr = sStr + cSep & .tbiRightImage.w                       '36
        sStr = sStr + cSep & .tbiRightImage.h                       '37
        
        sStr = sStr + cSep & .hBackColor                            '38
        sStr = sStr + cSep & .hTextColor                            '39
        sStr = sStr + cSep & .nBackColor                            '40
        sStr = sStr + cSep & .nTextColor                            '41
        
        sStr = sStr + cSep & Abs(.disabled)                         '42
        
        sStr = sStr + cSep & .caption                               '43
        sStr = sStr + cSep & Abs(.IncludeInToolbar)                 '44
        sStr = sStr + cSep & .ToolbarIndex                          '45
        sStr = sStr + cSep & .WinStatus                             '46
        
        sStr = sStr + cSep & .Actions.onclick.Type                  '47
        sStr = sStr + cSep & .Actions.onclick.url                   '48
        sStr = sStr + cSep & .Actions.onclick.TargetFrame           '49
        sStr = sStr + cSep & .Actions.onclick.TargetMenu            '50
        
        sStr = sStr + cSep & .Actions.onmouseover.Type              '51
        sStr = sStr + cSep & .Actions.onmouseover.url               '52
        sStr = sStr + cSep & .Actions.onmouseover.TargetFrame       '53
        sStr = sStr + cSep & .Actions.onmouseover.TargetMenu        '54
        
        sStr = sStr + cSep & .Actions.OnDoubleClick.Type            '55
        sStr = sStr + cSep & .Actions.OnDoubleClick.url             '56
        sStr = sStr + cSep & .Actions.OnDoubleClick.TargetFrame     '57
        sStr = sStr + cSep & .Actions.OnDoubleClick.TargetMenu      '58
        
        sStr = sStr + cSep & .CmdsMarginX                           '59
        sStr = sStr + cSep & .CmdsMarginY                           '60
        
        sStr = sStr + cSep & .ContentsMarginH                       '61
        sStr = sStr + cSep & .ContentsMarginV                       '62
        
        sStr = sStr + cSep & Replace(.Actions.onclick.WindowOpenParams, cSep, "|")       '63
        sStr = sStr + cSep & Replace(.Actions.onmouseover.WindowOpenParams, cSep, "|")   '64
        sStr = sStr + cSep & Replace(.Actions.OnDoubleClick.WindowOpenParams, cSep, "|") '65
        
        sStr = sStr + cSep & .HSImage                               '66
        
        sStr = sStr + cSep & Abs(.IsContext)                        '67
        sStr = sStr + cSep + .Sound.onmouseover                     '68
        sStr = sStr + cSep + .Sound.onclick                         '69
        
        sStr = sStr + cSep & .BorderStyle                           '70 Abs(.CmdsFXUseColor)
        sStr = sStr + cSep & .CmdsFXnColor                          '71
        sStr = sStr + cSep & .CmdsFXhColor                          '72
        
        sStr = sStr + cSep & .fHeight                               '73
        
        sStr = sStr + cSep & Abs(.IsTemplate)                       '74
        
        sStr = sStr + cSep & .Corners.bottomCorner                  '75
        sStr = sStr + cSep & .Corners.topCorner                     '76
        
        sStr = sStr + cSep & .AlignmentStyle                        '77
        
        sStr = sStr + cSep & .CornersImages.gcTopLeft               '78
        sStr = sStr + cSep & .CornersImages.gcTopCenter             '79
        sStr = sStr + cSep & .CornersImages.gcTopRight              '80
        sStr = sStr + cSep & .CornersImages.gcLeft                  '81
        sStr = sStr + cSep & .CornersImages.gcRight                 '82
        sStr = sStr + cSep & .CornersImages.gcBottomLeft            '83
        sStr = sStr + cSep & .CornersImages.gcBottomCenter          '84
        sStr = sStr + cSep & .CornersImages.gcBottomRight           '85
        
        sStr = sStr + cSep & .scrolling.maxHeight                   '86
        sStr = sStr + cSep & .scrolling.nColor                      '87
        sStr = sStr + cSep & .scrolling.hColor                      '88
        sStr = sStr + cSep & .scrolling.DnImage.NormalImage         '89
        sStr = sStr + cSep & .scrolling.DnImage.HoverImage          '90
        sStr = sStr + cSep & .scrolling.DnImage.w                   '91
        sStr = sStr + cSep & .scrolling.DnImage.h                   '92
        sStr = sStr + cSep & .scrolling.UpImage.NormalImage         '93
        sStr = sStr + cSep & .scrolling.UpImage.HoverImage          '94
        sStr = sStr + cSep & .scrolling.margin                      '95
        
        sStr = sStr + cSep & .DropShadowSize                        '96
        
        sStr = sStr + cSep & .tbiBackImage.NormalImage              '97
        sStr = sStr + cSep & .tbiBackImage.HoverImage               '98
        
        sStr = sStr + cSep & Abs(.scrolling.onmouseover)            '99
        sStr = sStr + cSep & .scrolling.FXhColor                    '100
        sStr = sStr + cSep & .scrolling.FXnColor                    '101
        sStr = sStr + cSep & .scrolling.FXNormal                    '102
        sStr = sStr + cSep & .scrolling.FXOver                      '103
        sStr = sStr + cSep & .scrolling.FXSize                      '104
        
        sStr = sStr + cSep & Abs(.Compile)                          '105
        
        sStr = sStr + cSep & Abs(.tbiBackImage.Tile)                '106
        sStr = sStr + cSep & Abs(.tbiBackImage.AllowCrop)           '107
        sStr = sStr + cSep & .tbiBackImage.w                        '108
        sStr = sStr + cSep & .tbiBackImage.h                        '109
        
        sStr = sStr + cSep & Abs(.DefNormalFont.FontShadow.Enabled1)                        '110
        sStr = sStr + cSep & .DefNormalFont.FontShadow.Color1                        '111
        sStr = sStr + cSep & .DefNormalFont.FontShadow.OffsetX1                        '112
        sStr = sStr + cSep & .DefNormalFont.FontShadow.OffsetY1                        '113
        sStr = sStr + cSep & .DefNormalFont.FontShadow.Blur1                        '114
        sStr = sStr + cSep & Abs(.DefNormalFont.FontShadow.Enabled2)                        '115
        sStr = sStr + cSep & .DefNormalFont.FontShadow.Color2                        '116
        sStr = sStr + cSep & .DefNormalFont.FontShadow.OffsetX2                        '117
        sStr = sStr + cSep & .DefNormalFont.FontShadow.OffsetY2                        '118
        sStr = sStr + cSep & .DefNormalFont.FontShadow.Blur2                        '119
        
        sStr = sStr + cSep & Abs(.DefHoverFont.FontShadow.Enabled1)                        '120
        sStr = sStr + cSep & .DefHoverFont.FontShadow.Color1                        '121
        sStr = sStr + cSep & .DefHoverFont.FontShadow.OffsetX1                        '122
        sStr = sStr + cSep & .DefHoverFont.FontShadow.OffsetY1                        '123
        sStr = sStr + cSep & .DefHoverFont.FontShadow.Blur1                        '124
        sStr = sStr + cSep & Abs(.DefHoverFont.FontShadow.Enabled2)                        '125
        sStr = sStr + cSep & .DefHoverFont.FontShadow.Color2                        '126
        sStr = sStr + cSep & .DefHoverFont.FontShadow.OffsetX2                        '127
        sStr = sStr + cSep & .DefHoverFont.FontShadow.OffsetY2                        '128
        sStr = sStr + cSep & .DefHoverFont.FontShadow.Blur2                        '129
        
        sStr = sStr + cSep & .tbiLeftImage.margin              '130
        sStr = sStr + cSep & .tbiRightImage.margin              '131
        
        sStr = sStr + cSep & .Radius.TopLeft                                    '132
        sStr = sStr + cSep & .Radius.TopRight                                    '133
        sStr = sStr + cSep & .Radius.BottomLeft                                    '134
        sStr = sStr + cSep & .Radius.BottomRight                                    '135
        
        sStr = sStr + cSep & .tbiRadius.TopLeft                                    '136
        sStr = sStr + cSep & .tbiRadius.TopRight                                    '137
        sStr = sStr + cSep & .tbiRadius.BottomLeft                                    '138
        sStr = sStr + cSep & .tbiRadius.BottomRight                                    '139
    End With
    
    GetGrpParams = sStr
    
End Function

Public Function GetCmdParams(cmd As MenuCmd) As String

    Dim sStr As String

    With cmd
        sStr = .Name                                                                ' 1
        sStr = sStr + cSep & .caption                                               ' 2
        sStr = sStr + cSep & .hBackColor                                            ' 3
        sStr = sStr + cSep & .hTextColor                                            ' 4
        sStr = sStr + cSep & .nBackColor                                            ' 5
        sStr = sStr + cSep & .nTextColor                                            ' 6
        sStr = sStr + cSep & .HoverFont.FontName                                    ' 7
        sStr = sStr + cSep & .HoverFont.FontSize                                    ' 8
        sStr = sStr + cSep & Abs(.HoverFont.FontBold)                               ' 9
        sStr = sStr + cSep & Abs(.HoverFont.FontItalic)                             '10
        sStr = sStr + cSep & Abs(.HoverFont.FontUnderline)                          '11
        sStr = sStr + cSep & .NormalFont.FontName                                   '12
        sStr = sStr + cSep & .NormalFont.FontSize                                   '13
        sStr = sStr + cSep & Abs(.NormalFont.FontBold)                              '14
        sStr = sStr + cSep & Abs(.NormalFont.FontItalic)                            '15
        sStr = sStr + cSep & Abs(.NormalFont.FontUnderline)                         '16
        sStr = sStr + cSep & vbNullString                                                     '17 (.URL)
        sStr = sStr + cSep & .iCursor.cType                                         '18
        sStr = sStr + cSep & .parent                                                '19
        sStr = sStr + cSep & .iCursor.CFile                                         '20 (.Trigger)
        sStr = sStr + cSep & .WinStatus                                             '21
        sStr = sStr + cSep & 0                                                      '22 (Abs(.ISCascade))
        sStr = sStr + cSep & 0                                                      '23 (.TargetMenu)
        sStr = sStr + cSep & .LeftImage.NormalImage                                 '24
        sStr = sStr + cSep & .LeftImage.HoverImage                                  '25
        sStr = sStr + cSep & 0                                                      '26 (Compat Old Versions)
        sStr = sStr + cSep & .LeftImage.w                                           '27
        sStr = sStr + cSep & .LeftImage.h                                           '28
        sStr = sStr + cSep & .Alignment                                             '29
        sStr = sStr + cSep & vbNullString                                                     '30 (Compat with NavIO)
        'sStr = sStr + cSep & .TargetFrame
        sStr = sStr + cSep & .Actions.onclick.TargetMenuAlignment                    '31 (Compat with 2.x)
        sStr = sStr + cSep & .Actions.onmouseover.TargetMenuAlignment                '32 (Compat with NavIO)
        sStr = sStr + cSep & .Actions.OnDoubleClick.TargetMenuAlignment              '33 (Compat with NavIO)
        sStr = sStr + cSep & Abs(.disabled)                                         '34
        sStr = sStr + cSep & .RightImage.NormalImage                                '35
        sStr = sStr + cSep & .RightImage.HoverImage                                 '36
        sStr = sStr + cSep & .RightImage.w                                          '37
        sStr = sStr + cSep & .RightImage.h                                          '38

        sStr = sStr + cSep & .Actions.onclick.Type                                  '39
        sStr = sStr + cSep & .Actions.onclick.url                                   '40
        sStr = sStr + cSep & .Actions.onclick.TargetFrame                           '41
        sStr = sStr + cSep & .Actions.onclick.TargetMenu                            '42
        
        sStr = sStr + cSep & .Actions.onmouseover.Type                              '43
        sStr = sStr + cSep & .Actions.onmouseover.url                               '44
        sStr = sStr + cSep & .Actions.onmouseover.TargetFrame                       '45
        sStr = sStr + cSep & .Actions.onmouseover.TargetMenu                        '46

        sStr = sStr + cSep & .Actions.OnDoubleClick.Type                            '47
        sStr = sStr + cSep & .Actions.OnDoubleClick.url                             '48
        sStr = sStr + cSep & .Actions.OnDoubleClick.TargetFrame                     '49
        sStr = sStr + cSep & .Actions.OnDoubleClick.TargetMenu                      '50
        
        sStr = sStr + cSep & .BackImage.NormalImage                                 '51
        sStr = sStr + cSep & .BackImage.HoverImage                                  '52
        
        sStr = sStr + cSep & Replace(.Actions.onclick.WindowOpenParams, cSep, "|")       '53
        sStr = sStr + cSep & Replace(.Actions.onmouseover.WindowOpenParams, cSep, "|")   '54
        sStr = sStr + cSep & Replace(.Actions.OnDoubleClick.WindowOpenParams, cSep, "|") '55
        
        sStr = sStr + cSep & .Sound.onmouseover                                     '56
        sStr = sStr + cSep & .Sound.onclick                                         '57
        
        sStr = sStr + cSep & .SeparatorPercent                                      '58
        
        sStr = sStr + cSep & .CmdsFXhColor                                      '59
        sStr = sStr + cSep & .CmdsFXnColor                                    '60
        sStr = sStr + cSep & .CmdsFXNormal                                      '61
        sStr = sStr + cSep & .CmdsFXOver                                      '62
        sStr = sStr + cSep & .CmdsFXSize                                      '63
        sStr = sStr + cSep & .CmdsMarginX                                      '64
        sStr = sStr + cSep & .CmdsMarginY                                      '65
        
        sStr = sStr + cSep & Abs(.Compile)                                      '66
        
        sStr = sStr + cSep & Abs(.BackImage.Tile)                                            '67
        sStr = sStr + cSep & Abs(.BackImage.AllowCrop)                                '68
        sStr = sStr + cSep & .BackImage.w                                                 '69
        sStr = sStr + cSep & .BackImage.h                                                '70
        
        sStr = sStr + cSep & Abs(.NormalFont.FontShadow.Enabled1)                        '71
        sStr = sStr + cSep & .NormalFont.FontShadow.Color1                        '72
        sStr = sStr + cSep & .NormalFont.FontShadow.OffsetX1                        '73
        sStr = sStr + cSep & .NormalFont.FontShadow.OffsetY1                        '74
        sStr = sStr + cSep & .NormalFont.FontShadow.Blur1                        '75
        sStr = sStr + cSep & Abs(.NormalFont.FontShadow.Enabled2)                        '76
        sStr = sStr + cSep & .NormalFont.FontShadow.Color2                        '77
        sStr = sStr + cSep & .NormalFont.FontShadow.OffsetX2                        '78
        sStr = sStr + cSep & .NormalFont.FontShadow.OffsetY2                        '79
        sStr = sStr + cSep & .NormalFont.FontShadow.Blur2                        '80
        
        sStr = sStr + cSep & Abs(.HoverFont.FontShadow.Enabled1)                        '81
        sStr = sStr + cSep & .HoverFont.FontShadow.Color1                        '82
        sStr = sStr + cSep & .HoverFont.FontShadow.OffsetX1                        '83
        sStr = sStr + cSep & .HoverFont.FontShadow.OffsetY1                        '84
        sStr = sStr + cSep & .HoverFont.FontShadow.Blur1                        '85
        sStr = sStr + cSep & Abs(.HoverFont.FontShadow.Enabled2)                        '86
        sStr = sStr + cSep & .HoverFont.FontShadow.Color2                        '87
        sStr = sStr + cSep & .HoverFont.FontShadow.OffsetX2                        '88
        sStr = sStr + cSep & .HoverFont.FontShadow.OffsetY2                        '89
        sStr = sStr + cSep & .HoverFont.FontShadow.Blur2                        '90
        
        sStr = sStr + cSep & .LeftImage.margin                                 '91
        sStr = sStr + cSep & .RightImage.margin                                  '92
        
        sStr = sStr + cSep & .Radius.TopLeft                                    '93
        sStr = sStr + cSep & .Radius.TopRight                                    '94
        sStr = sStr + cSep & .Radius.BottomLeft                                    '95
        sStr = sStr + cSep & .Radius.BottomRight                                    '96
    End With
    
    GetCmdParams = sStr

End Function

Public Sub DisableCloseButton(frm As Form)

    Dim hMenu As Long
    Dim menuItemCount As Long
    
    'Obtain the handle to the form's system menu
    hMenu = GetSystemMenu(frm.hwnd, 0)
    
    If hMenu Then
        'Obtain the number of items in the menu
         menuItemCount = GetMenuItemCount(hMenu)
        
        'Remove the system menu Close menu item.
        'The menu item is 0-based, so the last
        'item on the menu is menuItemCount - 1
         Call RemoveMenu(hMenu, menuItemCount - 1, MF_REMOVE Or MF_BYPOSITION)
        
        'Remove the system menu separator line
         Call RemoveMenu(hMenu, menuItemCount - 2, MF_REMOVE Or MF_BYPOSITION)
        
        'Force a redraw of the menu. This
        'refreshes the titlebar, dimming the X
         Call DrawMenuBar(frm.hwnd)
    End If

End Sub

Public Sub CoolListView(lv As ListView)

    Dim i As Long
    
    With lv
        If .View = lvwReport Then
            For i = 0 To .ColumnHeaders.Count - 1
                SendMessageLong .hwnd, LVM_SETCOLUMNWIDTH, i, ByVal LVSCW_AUTOSIZE_USEHEADER
            Next i
        End If
    End With

End Sub

Public Sub SelAll(txtCtrl As TextBox)

    With txtCtrl
        .SelStart = 0
        .SelLength = Len(.Text)
    End With

End Sub

Public Function RemoveLoaderCode(sCode As String, Optional FileName As String, Optional YahooSiteBuilder As Boolean = False) As String
    
    Dim p1 As Long
    Dim p2 As Long
    Dim l As Long
    Dim pCode As String
    Dim tCode As String
    Dim Ok2Remove As Boolean
    
    On Error GoTo AbortFunction
    
    ' Fix bug in previous versions
    pCode = Replace(sCode, "<!-- DHTML Menu Builder Loader Codee END -->", LoaderCodeEND)
    
    Do
        l = Len(pCode)
        If InStr(pCode, LoaderCodeSTART) > 0 And InStr(pCode, LoaderCodeEND) > 0 Then
            p1 = InStr(pCode, LoaderCodeSTART) - 1
            p2 = InStr(pCode, LoaderCodeEND) + Len(LoaderCodeEND)
            
            tCode = LCase(Mid(pCode, p1 + 1, p2 - p1))
            If CountOccurrences(tCode, "<p") > 0 Or _
                CountOccurrences(tCode, "<table") > 0 Or _
                CountOccurrences(tCode, "<br") > 0 Or _
                CountOccurrences(tCode, "<layer") > 0 Or _
                CountOccurrences(tCode, "<div") > 1 Or _
                CountOccurrences(tCode, "<span") > 0 Or _
                CountOccurrences(tCode, "<img") > 2 Or _
                CountOccurrences(tCode, "<script") > 2 Then
                
                If LenB(FileName) = 0 Then
                    Ok2Remove = True
                Else
                    If MsgBox("The Loader Code in the '" + FileName + "' document appears to have been modified and contain custom HTML code. If you continue this code will be lost." + vbCrLf + vbCrLf + "Are you sure you want to continue and remove the loader code along with any extra custom code that may have been inserted inside the menus' loader code?", vbQuestion + vbYesNo, "Remove Loader Code Warning") = vbYes Then
                        Ok2Remove = True
                    Else
                        Ok2Remove = False
                    End If
                End If
            Else
                Ok2Remove = True
            End If
            
            If Ok2Remove Then
                pCode = Left(pCode, p1) + Mid(pCode, p2)
            Else
                GoTo AbortFunction
            End If
        Else
            pCode = Replace(pCode, LoaderCodeSTART, vbNullString)
            pCode = Replace(pCode, LoaderCodeEND, vbNullString)
        
            p1 = InStr(pCode, "menu.js""")
            p2 = InStr(pCode, "_frames.js""")
            If p1 = p2 And p1 = 0 Then
                p1 = InStr(pCode, Project.JSFileName + ".js""")
                p2 = InStr(pCode, Project.JSFileName + "_frames.js""")
            End If
            If p1 > 0 Or p2 > 0 Then
                If p1 = 0 Then p1 = p2
                If p2 < p1 And p2 <> 0 Then p1 = p2
                p1 = InStrRev(pCode, "<script ", p1, vbTextCompare)
                p2 = InStr(p1, pCode, "</script>", vbTextCompare) + Len("</script>")
                
                pCode = Left(pCode, p1 - 1) + Mid(pCode, p2)
            End If
        End If
    Loop While l <> Len(pCode)
    
    If YahooSiteBuilder Then
        pCode = Replace(pCode, "  <!--$begin pageHtmlAfter$--> <!--$end pageHtmlAfter$-->", "")
        ' ...just in case...
        pCode = Replace(pCode, "<!--$begin pageHtmlAfter$--> <!--$end pageHtmlAfter$-->", "")
        pCode = Replace(pCode, "<!--$begin pageHtmlAfter$--><!--$end pageHtmlAfter$-->", "")
    End If
    
    RemoveLoaderCode = pCode
    
    Exit Function
    
AbortFunction:
    RemoveLoaderCode = sCode

End Function

Private Function CountOccurrences(ByVal c As String, ByVal s As String) As Integer

    On Error Resume Next
    CountOccurrences = UBound(Split(c, s))

End Function

Public Function AttachLoaderCode(ByVal pCode As String, LoaderCode As String, Optional YahooSiteBuilder As Boolean = False) As String
    
    Dim p As Long
    
    If YahooSiteBuilder Then
        p = LocateYSBSpot(pCode)
    Else
        p = LocateBODYtag(pCode)
    End If
    If p Then
        pCode = Left$(pCode, InStr(p, pCode, ">")) + _
                LoaderCode + Mid$(pCode, InStr(p, pCode, ">") + 1)
    Else
        pCode = LoaderCode + pCode
    End If
    
    AttachLoaderCode = pCode

End Function

Private Function LocateYSBSpot(ByVal sCode As String) As Long

    Const ym As String = "<!--$begin exclude$-->"
    Dim p As Long
    
    p = InStrRev(sCode, ym, , vbTextCompare)
    If p = 0 Then
        LocateYSBSpot = LocateBODYtag(sCode)
    Else
        LocateYSBSpot = p - Len(ym) - 1
    End If

End Function

Private Function LocateBODYtag(ByVal sCode As String) As Long

    Dim p As Long
    
    p = InStr(1, sCode, "</head", vbTextCompare)
    If p = 0 Then p = 1
    
    LocateBODYtag = InStr(p, sCode, "<body", vbTextCompare)

End Function

Private Sub MergePicture(idx As Integer)

    #If ISCOMP = 0 Then
    With frmMain
        .picItemIcon2.Picture = LoadPicture()
        .picItemIcon2.Picture = .ilIcons.ListImages(idx).Picture
        BitBlt .picItemIcon.hDc, 0, 0, 16, 16, .picItemIcon2.hDc, 0, 0, &H8800C6
        .picItemIcon.Picture = .picItemIcon.Image
    End With
    #End If

End Sub

Public Function GenCmdIcon(c As Integer) As Integer

    Dim pName As String
    Dim iIdx As Integer
    Dim nImg As ListImage
    
    #If ISCOMP = 0 Then
    frmMain.picItemIcon.Picture = LoadPicture()
    frmMain.picItemIcon.BackColor = &H80000005
        With MenuCmds(c)
        If .disabled Then
            MergePicture frmMain.ilIcons.ListImages("Disabled").Index
            pName = "Disabled"
        Else
            With .Actions
                If .onclick.Type = atcNone And _
                    .OnDoubleClick.Type = atcNone And _
                    .onmouseover.Type = atcNone Then
                    MergePicture frmMain.ilIcons.ListImages("NoEvents").Index
                    pName = "NoEvents"
                Else
                    pName = vbNullString
                    Select Case .onclick.Type
                        Case atcNone
                        Case atcCascade
                            MergePicture frmMain.ilIcons.ListImages("ClickCascade").Index
                            pName = pName + "ClickCascade"
                        Case Else
                            MergePicture frmMain.ilIcons.ListImages("Click").Index
                            pName = pName + "Click"
                    End Select
                    Select Case .OnDoubleClick.Type
                        Case atcNone
                        Case atcCascade
                            MergePicture frmMain.ilIcons.ListImages("DoubleClickCascade").Index
                            pName = pName + "DoubleClickCascade"
                        Case Else
                            MergePicture frmMain.ilIcons.ListImages("DoubleClick").Index
                            pName = pName + "DoubleClick"
                    End Select
                    Select Case .onmouseover.Type
                        Case atcNone
                        Case atcCascade
                            MergePicture frmMain.ilIcons.ListImages("OverCascade").Index
                            pName = pName + "OverCascade"
                        Case Else
                            MergePicture frmMain.ilIcons.ListImages("Over").Index
                            pName = pName + "Over"
                    End Select
                End If
            End With
        End If
    End With
    
    On Error Resume Next
    iIdx = frmMain.ilIcons.ListImages(pName).Index
    If iIdx = 0 Then
        Set nImg = frmMain.ilIcons.ListImages.Add(, pName, frmMain.picItemIcon.Picture)
        iIdx = nImg.Index
    End If
    
    GenCmdIcon = iIdx
    #End If
    
End Function

Public Function GenGrpIcon(g As Integer) As Integer

    #If ISCOMP = 0 Then

    Dim pName As String
    Dim iIdx As Integer
    Dim nImg As ListImage
    Dim prefix As String
    
    If MemberOf(g) = 0 Then
        prefix = "G"
    Else
        prefix = "T"
    End If
    'prefix = "G"
    
    frmMain.picItemIcon.Picture = LoadPicture()
    frmMain.picItemIcon.BackColor = &H80000005
    With MenuGrps(g)
        If .disabled Then
            MergePicture frmMain.ilIcons.ListImages("GDisabled").Index
            pName = "GDisabled"
        Else
            With .Actions
                If .onclick.Type = atcNone And _
                    .OnDoubleClick.Type = atcNone And _
                    .onmouseover.Type = atcNone Then
                    MergePicture frmMain.ilIcons.ListImages(prefix + "NoEvents").Index
                    pName = prefix + "NoEvents"
                Else
                    pName = vbNullString
                    Select Case .onclick.Type
                        Case atcNone
                        Case atcCascade
                            MergePicture frmMain.ilIcons.ListImages(prefix + "ClickCascade").Index
                            pName = pName + prefix + "ClickCascade"
                        Case Else
                            MergePicture frmMain.ilIcons.ListImages(prefix + "Click").Index
                            pName = pName + prefix + "Click"
                    End Select
                    Select Case .OnDoubleClick.Type
                        Case atcNone
                        Case atcCascade
                            MergePicture frmMain.ilIcons.ListImages(prefix + "DoubleClickCascade").Index
                            pName = pName + prefix + "DoubleClickCascade"
                        Case Else
                            MergePicture frmMain.ilIcons.ListImages(prefix + "DoubleClick").Index
                            pName = pName + prefix + "DoubleClick"
                    End Select
                    Select Case .onmouseover.Type
                        Case atcNone
                        Case atcCascade
                            MergePicture frmMain.ilIcons.ListImages(prefix + "OverCascade").Index
                            pName = pName + prefix + "OverCascade"
                        Case Else
                            MergePicture frmMain.ilIcons.ListImages(prefix + "Over").Index
                            pName = pName + prefix + "Over"
                    End Select
                End If
            End With
        End If
    End With
    
    On Error Resume Next
    iIdx = frmMain.ilIcons.ListImages(pName).Index
    If iIdx = 0 Then
        Set nImg = frmMain.ilIcons.ListImages.Add(, pName, frmMain.picItemIcon.Picture)
        iIdx = nImg.Index
    End If
    
    GenGrpIcon = iIdx
    #End If
    
End Function

Public Sub LoadLocalizedStrings()

    Dim lFile As String
    Dim lInfo As String
    Dim tmp() As String
    Dim i As Integer
    
    On Error GoTo LoadLocalizedStrings_Error

    ReDim LocalizedStr(0)
    ReDim engLocalizedStr(0)
    
    If LCase(Dir(AppPath + "lang\" + Preferences.language)) = LCase(Preferences.language) Then
        lFile = AppPath + "lang\" + Preferences.language
    Else
        lFile = AppPath + "lang\eng"
    End If
    
    If FileExists(lFile) Then
        lInfo = LoadFile(lFile)
        tmp = Split(lInfo, vbCrLf)
        
        For i = 100 To UBound(tmp)
            ReDim Preserve LocalizedStr(i + 1)
            LocalizedStr(i + 1) = tmp(i)
        Next i
    End If
    
    lFile = AppPath + "lang\eng"
    lInfo = LoadFile(lFile)
    tmp = Split(lInfo, vbCrLf)
    
    For i = 2 To UBound(tmp)
        ReDim Preserve engLocalizedStr(i + 1)
        engLocalizedStr(i + 1) = tmp(i)
    Next i

    On Error GoTo 0
    Exit Sub

LoadLocalizedStrings_Error:

    MsgBox "Error " & Err.number & " in line " & Erl & ": " & vbCrLf & Err.Description & " in Module.DMBGlobals.LoadLocalizedStrings"

End Sub

Public Function GetLocalizedStr(idx As Integer) As String

    Dim l As Integer
    
    On Error Resume Next
    
    If idx > UBound(LocalizedStr) Then
        If idx > UBound(engLocalizedStr) Then
            GetLocalizedStr = vbNullString
        Else
            GetLocalizedStr = engLocalizedStr(idx)
        End If
    Else
        If LenB(LocalizedStr(idx)) = 0 Then
            GetLocalizedStr = engLocalizedStr(idx)
            Exit Function
        End If
        
        GetLocalizedStr = LocalizedStr(idx)
    End If
    
    If InStr(GetLocalizedStr, "%%EMPTY-LINE%%") > 0 Then GetLocalizedStr = ""
    GetLocalizedStr = Replace(GetLocalizedStr, "%%CRLF%%", vbCrLf)
    
    While InStr(GetLocalizedStr, "%%LINE=")
        l = Val(Mid(GetLocalizedStr, InStr(GetLocalizedStr, "%%LINE=") + 7, 3))
        GetLocalizedStr = Left(GetLocalizedStr, InStr(GetLocalizedStr, "%%LINE=") - 1) + _
                            GetLocalizedStr(l) + _
                            Mid(GetLocalizedStr, InStr(GetLocalizedStr, "%%LINE=") + 12)
    Wend

End Function

Public Sub FixContolsWidth(frm As Form)

    Dim ctrl As Control
    Dim w As Integer
    
    For Each ctrl In frm.Controls
        If (TypeOf ctrl Is CheckBox) Or _
            (TypeOf ctrl Is OptionButton) Then
            If LenB(ctrl.caption) <> 0 Then
                w = GetTextSize(Unicode2xUNI(ctrl.caption), , ctrl)(1)
                ctrl.Width = w + 315 + IIf(ctrl.FontItalic, 60, 0)
            End If
        End If
    Next ctrl
    
End Sub

Public Function IconIndex(tagName As String) As Integer

    Dim img As ListImage
    
    For Each img In frmMain.ilIcons.ListImages
        If img.tag = tagName Then
            IconIndex = img.Index
            Exit Function
        End If
    Next img
    
    MsgBox "Unknown Image Resource: " + tagName

End Function

#If ISCOMP = 0 Or STANDALONE = 1 Then
Public Function DMBVersion() As String

    Dim ds As String
    Dim sp As String
    
    ds = GetDecimalSeparator
    DMBVersion = App.Major & ds & App.Minor & ds & Format(App.Revision, "000")
    
    sp = GetSetting("DMB", "RegInfo", "ServicePack", "")
    If LenB(sp) <> 0 Then DMBVersion = DMBVersion & " " & sp

End Function
#End If

Public Function FileIsHTML(ByVal FileName As String) As Boolean

    Dim p As Integer
    Dim ext As String
    
    FileName = LCase(FileName)
    
    p = InStrRev(FileName, ".") + 1
    If p = 1 Then Exit Function
    
    ext = Mid(FileName, p)
    FileIsHTML = (ext = "htm") Or (ext = "html")

End Function

Public Function ConfigTypeName(Config As ConfigDef) As String

    Select Case Config.Type
        Case ctcLocal
            ConfigTypeName = GetLocalizedStr(253)
        Case ctcRemote
            ConfigTypeName = GetLocalizedStr(254)
        Case ctcCDROM
            ConfigTypeName = "Relative" 'GetLocalizedStr(254)
        Case Else
            ConfigTypeName = GetLocalizedStr(472)
    End Select
    
End Function

Public Sub SaveLCFilesList(tv As TreeView, IsFramesLC As Boolean)

    Dim nItem As Node
    Dim ff As Integer
    Dim FileName As String
    
    On Error Resume Next
    MkDir AppPath + "LCLists"
    FileName = AppPath + "LCLists\" + GetFileName(Project.FileName) + "." + IIf(IsFramesLC, "f", "n") + "lc"
    
    ff = FreeFile
    Open FileName For Output As #ff
    For Each nItem In tv.Nodes
        If InStr(nItem.FullPath, ".") Then Print #ff, nItem.FullPath
    Next nItem
    Close ff

End Sub

Public Sub LoadLCFilesList(tv As TreeView, IsFramesLC As Boolean)

    Dim ff As Integer
    Dim FileName As String
    Dim sStr As String
    Dim f() As String
    Dim i As Integer
    Dim pNode As Node
    
    On Error Resume Next
    
    If LenB(Project.FileName) = 0 Then Exit Sub
    
    FileName = AppPath + "LCLists\" + GetFileName(Project.FileName) + "." + IIf(IsFramesLC, "f", "n") + "lc"
    
    If FileExists(FileName) Then
        ff = FreeFile
        Open FileName For Input As #ff
        Set pNode = tv.Nodes("[ROOT]")
        Do Until EOF(ff)
            Line Input #ff, sStr
            
            sStr = Replace(sStr, "Root Web", "[ROOT]")
            f = Split(sStr, "\")
            If UBound(f) = 1 Then
                Set pNode = tv.Nodes(1)
                i = 1
            Else
                Set pNode = tv.Nodes(1)
                For i = 1 To UBound(f) - 1
                    Err.Clear
                    Set pNode = tv.Nodes.Add(pNode.Index, tvwChild, "K" + f(i), f(i), 1)
                    If Err.number <> 0 Then
                        Set pNode = tv.Nodes("K" + f(i))
                    End If
                Next i
            End If
            tv.Nodes.Add pNode.Index, tvwChild, pNode.FullPath + "\" + "K" + f(i), f(i), 2
        Loop
        Close #ff
    End If
    
End Sub

Public Sub GetFramesInfo()

    With FramesInfo
        .Frames = ParseFrameset(.FileName)
        .IsValid = UBound(.Frames) > 0
    End With

End Sub

#If ISCOMP = 0 Then

Public Function IsHVVisible() As Boolean

    IsHVVisible = FindWindow("ThunderRT6FormDC", "DHTML Menu Builder Help")

End Function

Public Sub DrawColorBoxes(frm As Form)

    Dim ctrl As Control
    Dim sb As SmartButton
    Dim hDc As Long
    Dim r As RECT
    
    For Each ctrl In frm.Controls
        If TypeOf ctrl Is SmartButton Then
            Set sb = ctrl
            If (LenB(sb.caption) = 0 And sb.Visible) Then
                If sb.BackColor = frm.Point(1, 1) Then
                    hDc = GetDC(sb.Container.hwnd)
                    r.Left = sb.Left / Screen.TwipsPerPixelX - 1
                    r.Top = sb.Top / Screen.TwipsPerPixelY - 1
                    r.Right = (sb.Left + sb.Width) / Screen.TwipsPerPixelX + 1
                    r.bottom = (sb.Top + sb.Height) / Screen.TwipsPerPixelY + 1
                    DrawEdge hDc, r, BDR_RAISEDINNER Or BDR_RAISEDOUTER, BF_RECT Or BF_FLAT Or BF_SOFT
                    ReleaseDC sb.Container.hwnd, hDc
                End If
            ElseIf sb.Name = "sbApplyOptions" Then
                    hDc = GetDC(sb.Container.hwnd)
                    r.Left = sb.Left / Screen.TwipsPerPixelX - 1
                    r.Top = sb.Top / Screen.TwipsPerPixelY - 1
                    r.Right = (sb.Left + sb.Width) / Screen.TwipsPerPixelX + 1
                    r.bottom = (sb.Top + sb.Height) / Screen.TwipsPerPixelY + 1
                    'DrawEdge hDc, r, BDR_RAISEDINNER Or BDR_RAISEDOUTER, BF_RECT Or BF_FLAT Or BF_SOFT
                    DrawEdge hDc, r, BDR_SUNKENOUTER, BF_RECT Or BF_FLAT Or BF_SOFT
                    ReleaseDC sb.Container.hwnd, hDc
            End If
        End If
    Next ctrl

End Sub

Public Sub DisplayTip(Title As String, Message As String, Optional CanBeDisabled As Boolean = True)

    With TipsSys
        .TipTitle = Title
        .Tip = Message
        .CanDisable = CanBeDisabled
        .Show
    End With

End Sub

#End If

#End If

Public Sub tFont2Font(t As tFont, f As Font)

    f.Bold = t.FontBold
    f.Italic = t.FontItalic
    f.Name = t.FontName
    f.Size = t.FontSize
    f.Underline = t.FontUnderline

End Sub

Public Sub Font2tFont(f As Font, t As tFont)

    t.FontBold = f.Bold
    t.FontItalic = f.Italic
    t.FontName = f.Name
    t.FontSize = f.Size
    t.FontUnderline = f.Underline

End Sub

Public Sub SetColor(SelColor As Long, Control As Object)

    On Error Resume Next

    If SelColor <> -1 Then
        Control.Picture = LoadPicture()
        If SelColor = -2 Then
            Control.BackColor = vbBlack
            Control.PictureLayout = 6 'TopLeft
            frmMain.picRsc.Picture = LoadResPicture(103, vbResIcon)
            
            frmMain.picItemIcon.Width = Control.Width
            frmMain.picItemIcon.Height = Control.Height
            
            TileImage "picRsc", frmMain.picItemIcon
            
            Set frmMain.picItemIcon.Picture = frmMain.picItemIcon.Image
            Control.Picture = frmMain.picItemIcon.Picture
            
            frmMain.picItemIcon.Width = 240
            frmMain.picItemIcon.Height = 240
        Else
            Control.BackColor = SelColor
        End If
        Control.tag = SelColor
    End If

End Sub

Public Function IsInIDE() As Boolean

    On Error Resume Next
   
    Dim x As Long
    Debug.Assert Not TestIDE(x)
    IsInIDE = (x = 1) 'And False
      
End Function

Private Function TestIDE(ByRef x As Long) As Boolean
   
    x = 1
   
End Function

Public Function GetCurProjectVersion() As String

    Dim v() As String
    
    v = Split(GetFileVersion(Long2Short(AppPath + "dmb.exe")), ".")

    GetCurProjectVersion = v(0) + Format(v(1), "00") + Format(v(3), "000")

End Function

#If STANDALONE = 0 Then

Public Function GetSecuenceName(IsGroup As Boolean, prefix As String) As String

    Dim i As Integer
    Dim sName As String
    
    i = GetID
    
    If IsGroup Then
        Do
            sName = prefix & Format(i, "000")
            If ItemExists(sName, True) Then
                i = i + 1
            Else
                GetSecuenceName = sName
                Exit Function
            End If
        Loop
    Else
        Do
            sName = prefix & Format(i, "000")
            If ItemExists(sName, False) Then
                i = i + 1
            Else
                GetSecuenceName = sName
                Exit Function
            End If
        Loop
    End If
    
End Function

Public Function GetTBSecuenceName(prefix As String) As String

    Dim i As Integer
    Dim j As Integer
    Dim sName As String
    
    j = 0
    Do
        For i = 1 To UBound(Project.Toolbars)
            sName = prefix & Format(j, "000")
            If Project.Toolbars(i).Name = sName Then
                j = j + 1
                sName = ""
                Exit For
            End If
        Next i
        If LenB(sName) <> 0 Then Exit Do
    Loop
    
    GetTBSecuenceName = sName
    
End Function

Private Function ItemExists(Name As String, IsGroup As Boolean) As Boolean

    Dim i As Integer
    
    'If IsGroup Then
        For i = 1 To UBound(MenuGrps)
            If MenuGrps(i).Name = Name Then
                ItemExists = True
                Exit Function
            End If
        Next i
    'Else
        For i = 1 To UBound(MenuCmds)
            If MenuCmds(i).Name = Name Then
                ItemExists = True
                Exit Function
            End If
        Next i
    'End If
    
End Function

Public Sub ForceCommandsLinks2Local()

    If Project.UserConfigs(Project.DefaultConfig).Type = ctcRemote Then
        OriginalConfig = Project.DefaultConfig
        Project.DefaultConfig = GetConfigID(Project.UserConfigs(Project.DefaultConfig).LocalInfo4RemoteConfig)
        UpdateItemsLinks
    End If

End Sub

Public Sub RestoreCommandsLinks()

    If Project.UserConfigs(OriginalConfig).Type = ctcRemote Then
        Project.DefaultConfig = OriginalConfig
        UpdateItemsLinks
    End If

End Sub

#If ISCOMP = 0 Then

Public Function IsConnectedToInternet(Optional Silent As Boolean = False) As Boolean

    On Error Resume Next
    
    Dim dwFlags As Long
    Dim WebTest As Boolean
    
RetryIt:
    IsConnectedToInternet = IsOnline
    
    If Silent Then Exit Function

    If Not IsConnectedToInternet Then
        Select Case MsgBox(GetLocalizedStr(704), vbAbortRetryIgnore + vbInformation, GetLocalizedStr(703))
            Case vbAbort
                Exit Function
            Case vbIgnore
                IsConnectedToInternet = True
            Case vbRetry
                GoTo RetryIt
        End Select
    End If

End Function

Public Function CurEXEDate(Optional FileName As String = "") As Date

    If FileName = "" Then FileName = AppPath + "dmb.exe"
    
    If FileExists(FileName) Then
        CurEXEDate = CDate(Format(FileDateTime(FileName), "Short Date"))
    Else
        CurEXEDate = #11/19/1971#
    End If

End Function

Public Function SortListView(lv As ListView, SortType As Integer, lngIndex As Integer)

    On Error Resume Next
    
    ' Record the starting CPU time (milliseconds since boot-up)
    
    'Dim lngStart As Long
    'lngStart = GetTickCount
    
    lngIndex = lngIndex - 1
    
    ' Commence sorting
    
    With lv
    
        ' Display the hourglass cursor whilst sorting
        
        Dim lngCursor As Long
        lngCursor = .MousePointer
        .MousePointer = vbHourglass
        
        ' Prevent the ListView control from updating on screen -
        ' this is to hide the changes being made to the listitems
        ' and also to speed up the sort
        
        LockWindowUpdate .hwnd
        
        ' Check the data type of the column being sorted,
        ' and act accordingly
        
        Dim l As Long
        Dim strFormat As String
        Dim strData() As String
        
        Select Case SortType
        Case 0 ' DATE
        
            ' Sort by date.
            
            strFormat = "YYYYMMDDHhNnSs"
        
            ' Loop through the values in this column. Re-format
            ' the dates so as they can be sorted alphabetically,
            ' having already stored their visible values in the
            ' tag, along with the tag's original value
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .item(l).ListSubItems(lngIndex)
                            .tag = .Text & Chr$(0) & .tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    strFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .item(l)
                            .tag = .Text & Chr$(0) & .tag
                            If IsDate(.Text) Then
                                .Text = Format(CDate(.Text), _
                                                    strFormat)
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Sort the list alphabetically by this column
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = lngIndex
            .Sorted = True
            
            ' Restore the previous values to the 'cells' in this
            ' column of the list from the tags, and also restore
            ' the tags to their original values
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .item(l).ListSubItems(lngIndex)
                            strData = Split(.tag, Chr$(0))
                            .Text = strData(0)
                            .tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .item(l)
                            strData = Split(.tag, Chr$(0))
                            .Text = strData(0)
                            .tag = strData(1)
                        End With
                    Next l
                End If
            End With
            
        Case 1 ' NUMBER
        
            ' Sort Numerically
        
            strFormat = String(30, "0") & "." & String(30, "0")
        
            ' Loop through the values in this column. Re-format the values so as they
            ' can be sorted alphabetically, having already stored their visible
            ' values in the tag, along with the tag's original value
        
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .item(l).ListSubItems(lngIndex)
                            .tag = .Text & Chr$(0) & .tag
                            If InStr(.Text, " ") Then
                                Select Case LCase(Split(.Text, " ")(1))
                                    Case "kb"
                                        .Text = Val(Split(.Text, " ")(0)) * 1024
                                    Case "mb"
                                        .Text = Val(Split(.Text, " ")(0)) * (1024# * 1024#)
                                End Select
                            End If
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .item(l)
                            .tag = .Text & Chr$(0) & .tag
                            If InStr(.Text, " ") Then
                                Select Case LCase(Split(.Text, " ")(1))
                                    Case "kb"
                                        .Text = Val(Split(.Text, " ")(0)) * 1024
                                    Case "mb"
                                        .Text = Val(Split(.Text, " ")(0)) * (1024# * 1024#)
                                End Select
                            End If
                            If IsNumeric(.Text) Then
                                If CDbl(.Text) >= 0 Then
                                    .Text = Format(CDbl(.Text), _
                                        strFormat)
                                Else
                                    .Text = "&" & InvNumber( _
                                        Format(0 - CDbl(.Text), _
                                        strFormat))
                                End If
                            Else
                                .Text = ""
                            End If
                        End With
                    Next l
                End If
            End With
            
            ' Sort the list alphabetically by this column
            
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = lngIndex
            .Sorted = True
            
            ' Restore the previous values to the 'cells' in this
            ' column of the list from the tags, and also restore
            ' the tags to their original values
            
            With .ListItems
                If (lngIndex > 0) Then
                    For l = 1 To .Count
                        With .item(l).ListSubItems(lngIndex)
                            strData = Split(.tag, Chr$(0))
                            .Text = strData(0)
                            .tag = strData(1)
                        End With
                    Next l
                Else
                    For l = 1 To .Count
                        With .item(l)
                            strData = Split(.tag, Chr$(0))
                            .Text = strData(0)
                            .tag = strData(1)
                        End With
                    Next l
                End If
            End With
        
        Case Else ' STRING
            
            ' Sort alphabetically. This is the only sort provided
            ' by the MS ListView control (at this time), and as
            ' such we don't really need to do much here
        
            .SortOrder = (.SortOrder + 1) Mod 2
            .SortKey = lngIndex
            .Sorted = True
            
        End Select
    
        ' Unlock the list window so that the OCX can update it
        
        LockWindowUpdate 0&
        
        ' Restore the previous cursor
        
        .MousePointer = lngCursor
    
    End With
    
    ' Report time elapsed, in milliseconds
    
    'MsgBox "Time Elapsed = " & GetTickCount - lngStart & "ms"
    
End Function

Private Function InvNumber(ByVal number As String) As String
    Static i As Integer
    For i = 1 To Len(number)
        Select Case Mid$(number, i, 1)
        Case "-": Mid$(number, i, 1) = " "
        Case "0": Mid$(number, i, 1) = "9"
        Case "1": Mid$(number, i, 1) = "8"
        Case "2": Mid$(number, i, 1) = "7"
        Case "3": Mid$(number, i, 1) = "6"
        Case "4": Mid$(number, i, 1) = "5"
        Case "5": Mid$(number, i, 1) = "4"
        Case "6": Mid$(number, i, 1) = "3"
        Case "7": Mid$(number, i, 1) = "2"
        Case "8": Mid$(number, i, 1) = "1"
        Case "9": Mid$(number, i, 1) = "0"
        End Select
    Next
    InvNumber = number
End Function

Public Function NiceBytes(b As Long, Optional ShowDec As Boolean = True, Optional Reduce2KB As Boolean = False) As String

    Dim frmt As String
    Dim r As Integer
    
    If ShowDec Then
        r = 2
        frmt = "#,###.00"
    Else
        r = 0
        frmt = "#,###"
    End If

    Select Case b
        Case Is < 1024
            If Reduce2KB Then
                NiceBytes = Format(1, frmt) + " KB"
            Else
                NiceBytes = Format(b, frmt) & " b"
            End If
        Case Is < 1048576
            NiceBytes = Format(Round(b / 1024, r), frmt) + " KB"
        Case Else
            NiceBytes = Format(Round(b / (1024# * 1024#), r), frmt) + " MB"
    End Select

End Function

#End If

#End If

