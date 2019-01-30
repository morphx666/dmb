VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CEngine"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private cChar As String

Public Enum StatusOptionsConstants
    socStatus = 0
    socTooltip = 1
    socBoth = 2
End Enum

Public Enum SlashDir
    sdFwd = 0
    sdBack = 1
End Enum

Public Enum TriggerEventConstants
    ByHovering = 0
    ByClicking = 1
    ByDoubleClicking = 2
End Enum

Public Type BorderRadius
    TopLeft As Long
    TopRight As Long
    BottomLeft As Long
    BottomRight As Long
End Type

Public Enum CursorConstants
    iccDefault = 1
    iccCrosshair = 2
    iccHand = 3
    iccText = 4
    iccHelp = 5
    iccResizeE = 6
    iccResizeNE = 7
    iccResizeNW = 8
    iccResizeN = 9
    iccResizeSE = 10
    iccResizeSW = 11
    iccResizeS = 12
    iccResizeW = 13
    iccWait = 14
    iccCustom = 15
    iccResizeAll = 16
End Enum

Public Type tFontShadow
    Enabled1 As Boolean
    Color1 As Long
    OffsetX1 As Integer
    OffsetY1 As Integer
    Blur1 As Integer
    Enabled2 As Boolean
    Color2 As Long
    OffsetX2 As Integer
    OffsetY2 As Integer
    Blur2 As Integer
End Type

Public Type tCursor
    cType As CursorConstants
    cFile As String
End Type

Public Type tFont
    FontName As String
    FontBold As Boolean
    FontItalic As Boolean
    FontUnderline As Boolean
    FontSize As Integer
    FontShadow As tFontShadow
End Type

Public Type tImage
    NormalImage As String
    HoverImage As String
    w As Integer
    h As Integer
    Tile As Boolean
    AllowCrop As Boolean
    Margin As Integer
End Type

Public Enum TextAlignmentConstants
    tacLeft = 0
    tacRight = 1
    tacCenter = 2
End Enum

Public Enum CommandsFXConstants
    cfxcNone = 0
    cfxcRaised = 1
    cfxcSunken = 2
    cfxcDouble = 3
    cfxcBevel = 4
    cfxcDotted = 5
    cfxcDashed = 6
End Enum

Public Enum GroupAlignmentConstants
    gacBottomLeft = 0
    gacBottomRight = 1
    gacTopLeft = 2
    gacTopRight = 3
    gacLeftTop = 4
    gacLeftBottom = 5
    gacRightTop = 6
    gacRightBottom = 7
    gacTopCenter = 8
    gacRightCenter = 9
    gacBottomCenter = 10
    gacLeftCenter = 11
End Enum

Public Enum ActionTypeConstants
    atcNone = 0
    atcURL = 1
    atcCascade = 2
    atcNewWindow = 3
End Enum

Public Type Action
    url As String
    TargetFrame As String
    Type As ActionTypeConstants
    TargetMenu As Integer
    WindowOpenParams As String
    TargetMenuAlignment As GroupAlignmentConstants
End Type

Public Type ActionEvents
    OnMouseOver As Action
    OnClick As Action
    OnDoubleClick As Action
End Type

Public Type SoundDef
    OnMouseOver As String
    OnClick As String
End Type

Public Type MenuCmd
    Name As String
    Caption As String
    nTextColor As Long
    nBackColor As Long
    hTextColor As Long
    hBackColor As Long
    iCursor As tCursor
    Trigger As TriggerEventConstants
    Actions As ActionEvents
    WinStatus As String
    NormalFont As tFont
    HoverFont As tFont
    Parent As Integer
    LeftImage As tImage
    RightImage As tImage
    BackImage As tImage
    Alignment As TextAlignmentConstants
    Disabled As Boolean
    Sound As SoundDef
    xData As String
    SeparatorPercent As Integer
    
    CmdsFXNormal As CommandsFXConstants
    CmdsFXOver As CommandsFXConstants
    CmdsFXSize As Integer
    CmdsFXnColor As Long
    CmdsFXhColor As Long
    CmdsMarginX As Integer
    CmdsMarginY As Integer
    
    Radius As BorderRadius
    
    Compile As Boolean
End Type

Public Type CornersParams
    leftCorner As Long
    rightCorner As Long
    topCorner As Long
    bottomCorner As Long
End Type

Public Enum AlignmentStyleConstants
    [ascVertical]
    [ascHorizontal]
End Enum

Public Type GroupCorners
    gcTopLeft As String
    gcTopCenter As String
    gcTopRight As String
    gcLeft As String
    gcRight As String
    gcBottomLeft As String
    gcBottomCenter As String
    gcBottomRight As String
End Type

Public Type GrpScrollDef
    MaxHeight As Integer
    nColor As OLE_COLOR
    hColor As OLE_COLOR
    UpImage As tImage
    DnImage As tImage
    Margin As Integer
    OnMouseOver As Boolean
    FXNormal As CommandsFXConstants
    FXOver As CommandsFXConstants
    FXSize As Integer
    FXnColor As Long
    FXhColor As Long
End Type

Public Type MenuGrp
    Name As String
    Caption As String
    bColor As Long
    BorderStyle As CommandsFXConstants
    Corners As CornersParams
    FrameBorder As Integer
    DefNormalFont As tFont
    DefHoverFont As tFont
    Actions As ActionEvents
    x As Long
    y As Long
    Leading As Integer
    Image As String
    Alignment As GroupAlignmentConstants
    CmdsFXNormal As CommandsFXConstants
    CmdsFXOver As CommandsFXConstants
    CmdsFXSize As Integer
    'CmdsFXUseColor As Boolean
    CmdsFXnColor As Long
    CmdsFXhColor As Long
    CmdsMarginX As Integer
    CmdsMarginY As Integer
    ContentsMarginH As Integer
    ContentsMarginV As Integer
    DropShadowColor As Long
    DropShadowSize As Integer '1 to 255
    Transparency As Integer
    iCursor As tCursor
    Index As Integer
    fWidth As Integer
    fHeight As Integer
    CaptionAlignment As TextAlignmentConstants
    tbiLeftImage As tImage
    tbiRightImage As tImage
    tbiBackImage As tImage
    BackImage As tImage
    Disabled As Boolean
    nTextColor As Long
    nBackColor As Long
    hTextColor As Long
    hBackColor As Long
    WinStatus As String
    IncludeInToolbar As Boolean
    ToolbarIndex As Integer
    HSImage As String
    IsContext As Boolean
    Sound As SoundDef
    IsTemplate As Boolean
    AlignmentStyle As AlignmentStyleConstants
    xData As String
    CornersImages As GroupCorners
    Scrolling As GrpScrollDef
    Radius As BorderRadius
    tbiRadius As BorderRadius
    
    Compile As Boolean
End Type

Public Type Section
    Name As String
    params As String
    NewParams As String
    Code As String
    NewCode As String
    Support As String
    Description As String
    IsNew As Boolean
    ChangedFlag As Boolean
End Type

Public Type AddInParameter
    Name As String
    Description As String
    Value As String
    Default As String
    Required As Boolean
End Type

Public Enum TargetPlatformConstants
    tpcIE = 0
    tpcNS = 1
    tpcBoth = 2
End Enum

Public Enum CodeOptimizationConstants
    cocDEBUG = 0
    cocNormal = 1
    cocAggressive = 2
End Enum

Public Type AddInDef
    Name As String
    Description As String
End Type

Public Enum ToolbarAlignmentConstants
    tbacTopLeft = 0
    tbacTopCenter = 1
    tbacTopRight = 2
    tbacCenterLeft = 3
    tbacCenterCenter = 4
    tbacCenterRight = 5
    tbacBottomLeft = 6
    tbacBottomCenter = 7
    tbacBottomRight = 8
    tbacCustom = 9
    tbacAttached = 10
    tbacFree = 11
End Enum

Public Enum ToolbarStyleConstants
    tscHorizonal = 0
    tscVertical = 1
End Enum

Public Enum ToolbarSpanningConstants
    tscAuto = 0
    tscFull = 1
End Enum

Public Type ToolbarDef
    Name As String
    CreateToolbar As Boolean
    FollowHScroll As Boolean
    FollowVScroll As Boolean
    SmartScrolling As Boolean
    Alignment As ToolbarAlignmentConstants
    AttachTo As String
    AttachToAlignment As GroupAlignmentConstants
    AttachToAutoResize As Boolean
    Style As ToolbarStyleConstants
    Spanning As ToolbarSpanningConstants
    Border As Integer
    BorderColor As OLE_COLOR
    BorderStyle As CommandsFXConstants
    JustifyHotSpots As Boolean
    BackColor As OLE_COLOR
    Image As String
    CustX As Integer
    CustY As Integer
    Groups() As String
    Separation As Integer
    Width As Integer
    Height As Integer
    ContentsMarginH As Integer
    ContentsMarginV As Integer
    Condition As String
    Transparency As Integer
    DropShadowColor As Long
    DropShadowSize As Integer
    IsTemplate As Boolean
    
    OffsetH As Integer
    OffsetV As Integer
    
    Radius As BorderRadius
    
    Compile As Boolean
End Type

Public Type HotSpotEditorConfigDef
    HotSpotsFile As String
    MakeBackup As Boolean
End Type

Public Type FramesConfigDef
    UseFrames As Boolean
    FramesFile As String
End Type

Public Enum ConfigTypeContants
    [ctcLocal]
    [ctcRemote]
    [ctcCDROM]
End Enum

Public Type ConfigDef
    Name As String
    Description As String
    
    RootWeb As String
    CompiledPath As String
    ImagesPath As String

    OptmizePaths As Boolean
    
    HotSpotEditor As HotSpotEditorConfigDef
    Frames As FramesConfigDef
    
    Type As ConfigTypeContants
    
    LocalInfo4RemoteConfig As String
    FTP As String
End Type

'Public Type FTPConfigDef
'    FTPAddress As String
'    UserName As String
'    Password As String
'    ProxyAddress As String
'    ProxyPort As Long
'    RemoteInfo4FTP As String
'End Type

Public Type MenusOffsetDef
    RootMenusX As Integer
    RootMenusY As Integer
    SubMenusX As Integer
    SubMenusY As Integer
End Type

Public Enum exHTMLStylesConstants
    ascNone = 0
    ascProject = 1
    ascCSS = 2
End Enum

Public Enum exHTMLExpColPlacementConstants
    ecpcTop = 0
    ecpcBottom = 1
    ecpcBoth = 2
End Enum

Public Type ExportHTMLDef
    Title As String
    Description As String
    Style As exHTMLStylesConstants
    GroupClass As String
    CommandClass As String
    HTMLFileName As String
    ImagesPath As String
    CreateTree As Boolean
    CollapsedImage As String
    ExpandedImage As String
    NormalImage As String
    IconWidth As Integer
    IconHeight As Integer
    Identation As String
    CSSFile As String
    ExpItemsHaveLinks As Boolean
    SingleSelect As Boolean
    IncludeExpCol As Boolean
    ExpAllStr As String
    ColAllStr As String
    ExpColPlacement As exHTMLExpColPlacementConstants
    XHTMLCompliant As Boolean
End Type

Public Type ProjectDef
    Name As String
    FileName As String
    HasChanged As Boolean
    OriginalFileName As String
    
    FX As Integer
    UnfoldingSound As SoundDef
    CodeOptimization As CodeOptimizationConstants
    RemoveImageAutoPosCode As Boolean
    UseGZIP As Boolean
    
    AbsPath As String 'Obsolete
    GenDynAPI As Boolean
    
    UserConfigs() As ConfigDef
    DefaultConfig As Integer
    
    'FTP As FTPConfigDef
    AddIn As AddInDef
    Toolbar As ToolbarDef
    Toolbars() As ToolbarDef
    
    JSFileName As String
    
    Version As String
    
    MenusOffset As MenusOffsetDef
    
    FontSubstitutions As String
    
    AnimSpeed As Integer
    HideDelay As Integer
    SubMenusDelay As Integer
    RootMenusDelay As Integer
    SelChangeDelay As Integer
    
    DoFormsTweak As Boolean
    DWSupport As Boolean
    NS4ClipBug As Boolean
    OPHelperFunctions As Boolean
    ImageReadySupport As Boolean
    LotusDominoSupport As Boolean
    AutoSelFunction As Boolean
    KeyboardSupport As Boolean
    AutoScroll As GrpScrollDef
    SEOTweak As Boolean
    
    CompileIECode As Boolean
    CompileNSCode As Boolean
    CompilehRefFile As Boolean
    
    DXFilter As String
    BlinkEffect As Integer
    BlinkSpeed As Integer
    
    DOMCode As String
    DOMFramesCode As String
    NSCode As String
    NSFramesCode As String
    
    ExportHTMLParams As ExportHTMLDef
    
    SecondaryProjects() As String
    
    NodeExpStatus As String
    
    DBCSSupport As Boolean
    
    StatusTextDisplay As StatusOptionsConstants
    
    CustomOffsets As String
End Type

Public Enum DataOnClibboardConstants
    docCommand = 1
    docGroup = 2
    docSeparator = 3
    docToolbar = 4
End Enum

Public Type ClipboardDef
    CmdContents As MenuCmd
    GrpContents As MenuGrp
    TBContents As ToolbarDef
    ObjSrc As DataOnClibboardConstants
    CustomSel() As String
End Type

Public Type CFrame
    Name As String
    SrcFile As String
End Type

Public Type FramesInfoDef
    FileName As String
    Frames() As CFrame
    IsValid As Boolean
End Type

Public Enum InheritanceConstants
    [icDefault]
    [icFirst]
End Enum

Public Type ItemStyleDef
    Font As tFont
    Color As OLE_COLOR
End Type

Public Enum LinkVerificationConstants
    lvcDoNotVerify = 0
    lvcVerifyWhenOpening = &H100
    lvcVerifyWhenSaving = &H10
    lvcVerifyWhenCompiling = &H1
End Enum

Public Type LinkVerycationDef
    VerifyOptions As Long
    VerifyExternalLinks As Boolean
End Type

Public Type PrgPrefs
    OpenLastProject As Boolean
    AutoRecover As Boolean
    ShowNag As Boolean
    
    ShowWarningAddInEditor As Boolean
    ShowPPOnNewProject As Boolean
    ShowCleanPreview As Boolean
    
    SepHeight As Integer
    ImgSpace As Integer
    CommandsInheritance As InheritanceConstants
    GroupsInheritance As InheritanceConstants
    
    EnableUndoRedo As Boolean
    UseLivePreview As Boolean
    
    Language As String
    
    UseInstallMenus As Boolean
    UseMapView As Boolean
    LockProjects As Boolean
    UseEasyActions As Boolean
    
    CodePage As String
    
    ToolbarStyle As ItemStyleDef
    ToolbarItemStyle As ItemStyleDef
    CommandStyle As ItemStyleDef
    GroupStyle As ItemStyleDef
    BrokenLink As OLE_COLOR
    DisabledItem As OLE_COLOR
    NoCompileItem As OLE_COLOR
    
    VerifyLinksOptions As LinkVerycationDef
    AutoShowCompileReport As Boolean
    
    EnableUnicodeInput As Boolean
End Type

Private Type ExandableItemDef
    sStr As String
    Platform As TargetPlatformConstants
    CanExpandArray As Boolean
End Type
Private ExpandableItems() As ExandableItemDef
'Private Sounds() As String
Private leftCorner As Integer
Private topCorner As Integer
Private rightCorner As Integer
Private bottomCorner As Integer

Private tbBaseIndex As Integer
Private gBaseIndex As Integer
Private cBaseIndex As Integer
Private IsMultiProject As Boolean

Private mIsCompiling As Boolean

Public Function EngineVersion() As String

    Dim ds As String
    
    ds = GetDecimalSeparator

    EngineVersion = App.Major & ds & Format$(App.Minor, "00") & ds & Format$(App.Revision, "000")

End Function

Public Sub SetUI(Objects() As Object)

    Set frmProjProp = Objects(1)
    Set FloodPanel.PictureControl = Objects(2).PictureControl
    Set frmMain = Objects(3)

End Sub

Public Sub SetVars(Objects() As Variant)

    AppPath = Objects(1)
    If LenB(GenLicense) = 0 Then
        HelpFile = Objects(2)
        GenLicense = Objects(3)
    End If
    TempPath = Objects(4)
    cSep = Objects(5)
    nwdPar = Objects(6)
    StatesPath = Objects(7)

End Sub

Private Function CreateFlashCode(FileName As String, w As String, h As String, bColor As String, IsNS As Boolean) As String

    Dim sCode As String
    
    If IsNS Then
        sCode = "<layer" + _
                    " left=0" & _
                    " top=0" & _
                    " width=" & w & _
                    " height=" & h & _
                ">"
    Else
        sCode = "<span style=""position:absolute" + _
                "; top:0px" & _
                "; left:0px" & _
                "; width:" & w & IIf(InStr(w, "%"), "", "px") & _
                "; height:" & h & IIf(InStr(h, "%"), "", "px") & _
            ";"">"
    End If
    
    sCode = sCode + "<embed src=""" + ImgAbsPath + GetFileName(FileName) + """ quality=high " + _
                      AddConditionalAttribute("bcolor", bColor, True) + _
                      " width=" & w & " height=" & h & _
                      " WMode=Transparent type=""application/x-shockwave-flash"" pluginspage=""http://www.macromedia.com/shockwave/download/index.cgi?p1_prod_version=shockwaveflash""></embed>"
    
    sCode = sCode + IIf(IsNS, "</layer>", "</span>")
    
    CreateFlashCode = sCode

End Function

Public Property Get IsCompiling() As Boolean

    IsCompiling = mIsCompiling

End Property

Public Function CompileProject(mg() As MenuGrp, mc() As MenuCmd, PD As ProjectDef, Prefs As PrgPrefs, _
                                AddInParams() As AddInParameter, Optional IsForPreviewing As Boolean = False, _
                                Optional DoNotCompile As Boolean = False, Optional vPreviewPath As String, Optional DoNotSave As Boolean = False, _
                                Optional tbIndex As Integer = 0, Optional cmdIndex As Integer = 0, Optional grpIndex As Integer = 0, _
                                Optional IsLivePreview As Boolean = False) As Boolean

    Dim g As Integer
    Dim c As Integer
    Dim ieTable As String
    Dim nsTable As String
    Dim ieStyleCode As String
    Dim nsStyleCode As String
    Dim bStyleCode As String
    Dim hRefStatements() As String
    Dim hLayers() As String
    Dim nLayers() As String
    Dim lIdx As Integer
    Dim id As Integer
    Dim sid As Integer
    Dim cmdTop As Integer
    Dim CmdLayerName As String
    Dim cFX() As String
    Dim XYPos As String
    
    Dim fTotal As Integer
    Dim cc As Integer
    
    'Cache Variables
    Dim DivWidth As Integer
    Dim DivHeight As Integer
    Dim nTColor As String
    Dim hTColor As String
    Dim nBColor As String
    Dim hBColor As String
    Dim gBColor As String
    Dim scHeight As Integer
    Dim scWidth As Integer
    Dim nsEC As String
    Dim cName As String
    Dim nsImgTblNorm As String
    Dim nsImgTblHovr As String
    Dim nsEventCode() As String
    Dim dispTips As Boolean
    
    mIsCompiling = True
    
    Dim fnp As String
    fnp = TempPath + GetFileName(Project.FileName, True)
    
    CompileProject = False
    
    MenuGrps = mg
    MenuCmds = mc
    Project = PD
    params = AddInParams
    
    If IsForPreviewing Then
        tbBaseIndex = 0
        gBaseIndex = 0
        cBaseIndex = 0
        IsMultiProject = False
    Else
        tbBaseIndex = tbIndex
        gBaseIndex = grpIndex
        cBaseIndex = cmdIndex
        IsMultiProject = (tbBaseIndex > 0 Or cBaseIndex > 0)
        'IsMultiProject = False
    End If
    
    If UBound(MenuGrps) = 0 Then
        MsgBox "There is nothing to compile!", vbInformation + vbOKOnly, "Error Compiling"
        GoTo ExitSub
    End If
    
    If Not IsLivePreview Then frmMain.MousePointer = vbHourglass
    FloodPanel.IgnoreChanges = IsLivePreview
    
    Preferences = Prefs
    ThisConfig = Project.UserConfigs(Project.DefaultConfig)
    
    If IsForPreviewing Then
        JSAbsPath = ""
        ImgAbsPath = ""
        PreviewPath = vPreviewPath
    Else
        Select Case ThisConfig.Type
            Case ctcRemote
                If IsExternalLink(ThisConfig.CompiledPath) Then
                    JSAbsPath = ThisConfig.CompiledPath
                Else
                    JSAbsPath = SetSlashDir(ThisConfig.RootWeb + ThisConfig.CompiledPath, sdFwd)
                End If
                If IsExternalLink(ThisConfig.ImagesPath) Then
                    ImgAbsPath = ThisConfig.ImagesPath
                Else
                    ImgAbsPath = SetSlashDir(ThisConfig.RootWeb + ThisConfig.ImagesPath, sdFwd)
                End If
                If ThisConfig.OptmizePaths Then
                    JSAbsPath = Replace(Mid(JSAbsPath, InStr(InStr(JSAbsPath, "//") + 2, JSAbsPath, "/")), "//", "/")
                    ImgAbsPath = Replace(Mid(ImgAbsPath, InStr(InStr(ImgAbsPath, "//") + 2, ImgAbsPath, "/")), "//", "/")
                Else
                    JSAbsPath = Replace(JSAbsPath, "//", "/")
                    JSAbsPath = Replace(JSAbsPath, ":/", "://")
                    ImgAbsPath = Replace(ImgAbsPath, "//", "/")
                    ImgAbsPath = Replace(ImgAbsPath, ":/", "://")
                End If
            Case ctcLocal
                JSAbsPath = AddFileProtocol(ThisConfig.CompiledPath)
                ImgAbsPath = AddFileProtocol(ThisConfig.ImagesPath)
            Case ctcCDROM
                JSAbsPath = ""
                ImgAbsPath = "%%REL%%"
        End Select
        
        #If DEMO = 0 Then
        DblChk
        #End If
    End If
    
    JSAbsPath = EscapePath(JSAbsPath)
    ImgAbsPath = EscapePath(ImgAbsPath)
    
    FramesMode = ThisConfig.Frames.UseFrames
    
    If DoNotCompile Then GoTo ExitSub
    
    dispTips = Project.StatusTextDisplay = socBoth Or Project.StatusTextDisplay = socTooltip
    fTotal = UBound(MenuGrps) + UBound(MenuCmds)
    
    ieTable = ""
    nsTable = ""
    
    id = cBaseIndex
    lIdx = cBaseIndex

    'Erase Sounds: ReDim Sounds(0)
    'Add2SoundsArray Project.UnfoldingSound.OnMouseOver
    ReDim hRefStatements(1 To UBound(MenuGrps))
    
    If Not IsForPreviewing And Project.CodeOptimization = cocAggressive Then
        For g = 1 To UBound(MenuGrps)
            If Left(MenuGrps(g).Name, 9) = "SubGroup_" Then MenuGrps(g).Name = "g" & g
        Next g
    End If
    
    UsesDynAPI = False
    For g = 1 To UBound(MenuGrps)
        If MenuGrps(g).IsTemplate Then
            UsesDynAPI = True
            Exit For
        End If
    Next g
    
    sid = 999
    
    For g = 1 To UBound(MenuGrps) + Len(QueryValue(HKEY_CLASSES_ROOT, "*", "ContentType"))
        If MenuGrps(g).Compile Then
    
            'Add2SoundsArray MenuGrps(g).Sound.OnMouseOver
            'Add2SoundsArray MenuGrps(g).Sound.OnClick
        
            DivWidth = GetDivWidth(g)
            DivHeight = GetDivHeight(g)
            gBColor = GetRGB(MenuGrps(g).bColor, True)
            
            leftCorner = CornerSize(g, "left")
            rightCorner = CornerSize(g, "right")
            topCorner = CornerSize(g, "top")
            bottomCorner = CornerSize(g, "bottom")
        
            FloodPanel.Caption = "Compiling: " + NiceGrpCaption(g)
            
            'Generate the calling code to display the menus
            hRefStatements(g) = "# CALL TO MENU: " + MenuGrps(g).Name + vbCrLf + _
                                GetGroupEventCode(g, MenuGrps(g).HSImage, IsForPreviewing) + vbCrLf
            
            'Create the group's DIVs AND LAYERs
            ieTable = ieTable + "<div id=""" + MenuGrps(g).Name + _
                                """ " + IIf(UsesDynAPI, "fh=" & IIf(MenuGrps(g).fHeight < 0, Abs(MenuGrps(g).fHeight), 0) & " fw=" & IIf(MenuGrps(g).fWidth < 0, Abs(MenuGrps(g).fWidth), 0) & " cl=" & MenuGrps(g).AlignmentStyle & " ld=" & MenuGrps(g).Leading & " h=" & MenuGrps(g).ContentsMarginH & " v=" & MenuGrps(g).ContentsMarginV & " tc=" & topCorner + bottomCorner + MenuGrps(g).DropShadowSize & " lc=" & leftCorner + rightCorner + MenuGrps(g).DropShadowSize & " ", "") + "style=""position:absolute; top:0px; left:0px" + _
                                "; width:" & DivWidth + leftCorner + rightCorner + MenuGrps(g).DropShadowSize & "px" + _
                                "; height:" & DivHeight + topCorner + bottomCorner + MenuGrps(g).DropShadowSize & "px" & _
                                RaiusToCSS(MenuGrps(g).Radius) & _
                                " ;" + IIf(MenuGrps(g).Transparency > 0, MozillaAndSafariTrans(MenuGrps(g).Transparency), "") + _
                                "; visibility:hidden;" + _
                                """" + IIf(topCorner <> 0 Or bottomCorner <> 0, " tci=" & topCorner & " bci=" & bottomCorner, "") + ">" + vbCrLf + _
                                CreateCorners(g, DivWidth, DivHeight, False)
            ieTable = ieTable + "<div id=""" + MenuGrps(g).Name + "frmt"" style=""position:absolute" + _
                                "; top:" & topCorner & "px" + _
                                "; left:" & leftCorner & "px" + _
                                "; width:" & DivWidth & "px" + _
                                "; height:" & DivHeight & "px" & _
                                RaiusToCSS(MenuGrps(g).Radius) & _
                                " ;" + IIf(MenuGrps(g).Transparency > 0, MozillaAndSafariTrans(MenuGrps(g).Transparency), "") + _
                                " ;" + IIf(MenuGrps(g).DropShadowSize > 0, MozillaAndSafariShadow(MenuGrps(g).DropShadowSize, MenuGrps(g).DropShadowColor), "") + _
                                "; " + AddConditionalAttribute("background-color", gBColor, False) + _
                                "; " + IIf(Not IsFlash(MenuGrps(g).Image), AddConditionalAttribute("background-image", "url(" + ImgAbsPath + GetFileName(MenuGrps(g).Image) + ")", False), "") + _
                                IIf(MenuGrps(g).FrameBorder > 0, _
                                    IIf(MenuGrps(g).Corners.leftCorner <> -2, "; border-left:" + GetRGB(MenuGrps(g).Corners.leftCorner, True) + " " & MenuGrps(g).FrameBorder & "px " + GetBorderStyleName(MenuGrps(g).BorderStyle), "") + _
                                    IIf(MenuGrps(g).Corners.rightCorner <> -2, "; border-right:" + GetRGB(MenuGrps(g).Corners.rightCorner, True) & " " & MenuGrps(g).FrameBorder & "px " + GetBorderStyleName(MenuGrps(g).BorderStyle), "") + _
                                    IIf(MenuGrps(g).Corners.topCorner <> -2, "; border-top:" + GetRGB(MenuGrps(g).Corners.topCorner, True) & " " & MenuGrps(g).FrameBorder & "px " + GetBorderStyleName(MenuGrps(g).BorderStyle), "") + _
                                    IIf(MenuGrps(g).Corners.bottomCorner <> -2, "; border-bottom:" + GetRGB(MenuGrps(g).Corners.bottomCorner, True) & " " & MenuGrps(g).FrameBorder & "px " + GetBorderStyleName(MenuGrps(g).BorderStyle), "") + ";", "") + _
                                """>" + IIf(IsFlash(MenuGrps(g).Image), CreateFlashCode(MenuGrps(g).Image, CStr(DivWidth), CStr(DivHeight), GetRGB(MenuGrps(g).bColor, True), False), "") + vbCrLf
            
            If Project.CompileNSCode Then
            nsTable = nsTable + "<layer name=" + MenuGrps(g).Name & _
                                " top=0 left=0" & _
                                " width=" & DivWidth + leftCorner + rightCorner & _
                                " height=" & DivHeight + topCorner + bottomCorner & _
                                " z-index=1000" + IIf(leftCorner + topCorner + rightCorner + bottomCorner > 0, "", IIf((gBColor <> "-2") Or (LenB(MenuGrps(g).Image) <> 0), AddConditionalAttribute(" bgColor", GetRGB(MenuGrps(g).Corners.leftCorner), True), "")) + _
                                " visibility=hidden>" + vbCrLf + _
                                CreateCorners(g, DivWidth, DivHeight, True) + _
                                "<layer" + _
                                AddConditionalAttribute(" bgColor", gBColor, True) & _
                                IIf(Not IsFlash(MenuGrps(g).Image), AddConditionalAttribute(" background", ImgAbsPath + GetFileName(MenuGrps(g).Image), True), "") & _
                                " left=" & MenuGrps(g).FrameBorder + leftCorner & _
                                " top=" & MenuGrps(g).FrameBorder + topCorner & _
                                " width=" & DivWidth - CalcFrameBorder(g, ascHorizontal) & _
                                " height=" & DivHeight - CalcFrameBorder(g, ascVertical) & _
                                " z-index=1001>" + IIf(IsFlash(MenuGrps(g).Image), CreateFlashCode(MenuGrps(g).Image, DivWidth + leftCorner + rightCorner, DivHeight + topCorner + bottomCorner, GetRGB(MenuGrps(g).bColor), True), "") + vbCrLf
            End If
            
            Select Case MenuGrps(g).AlignmentStyle
                Case ascVertical
                    cmdTop = MenuGrps(g).ContentsMarginV
                Case ascHorizontal
                    cmdTop = MenuGrps(g).ContentsMarginH
            End Select
            
            For c = 1 To UBound(MenuCmds)
                If MenuCmds(c).Parent = g And MenuCmds(c).Compile Then
                    cFX = GetBorderFX(c, False)
                    cc = cc + 1
                    FloodPanel.Value = CInt((g + cc) / fTotal * 100)
                    
                    DoEvents
                
                    'Add2SoundsArray MenuCmds(c).Sound.OnMouseOver
                    'Add2SoundsArray MenuCmds(c).Sound.OnClick
                
                    nTColor = GetRGB(MenuCmds(c).nTextColor)
                    hTColor = GetRGB(MenuCmds(c).hTextColor)
                    nBColor = GetRGB(MenuCmds(c).nBackColor, True)
                    hBColor = GetRGB(MenuCmds(c).hBackColor, True)
                    Select Case MenuGrps(g).AlignmentStyle
                        Case ascVertical
                            scWidth = DivWidth - CalcFrameBorder(g, ascHorizontal) - 2 * MenuGrps(g).ContentsMarginH
                            scHeight = GetSingleCommandHeight(c)
                        Case ascHorizontal
                            scWidth = GetSingleCommandWidth(c)
                            scHeight = DivHeight - CalcFrameBorder(g, ascVertical) - 2 * MenuGrps(g).ContentsMarginV - 2 * MenuCmds(c).CmdsFXSize
                    End Select
                    cName = GetCursorName(MenuCmds(c).iCursor)
                    
                    lIdx = lIdx + 1
                    ReDim Preserve nLayers(lIdx)
                    ReDim Preserve hLayers(lIdx)
                    
                    If Project.CompileIECode Or IsForPreviewing Then
                        If MenuGrps(g).AlignmentStyle = ascVertical Then
                            XYPos = "; top:" & cmdTop & "px" & _
                                    "; left:" & MenuGrps(g).ContentsMarginH & "px"
                                    
                        Else
                            XYPos = "; top:" & MenuGrps(g).ContentsMarginV & "px" & _
                                    "; left:" & cmdTop & "px"
                                    
                        End If
                        If MenuCmds(c).Name = "[SEP]" Then
                            If MenuGrps(g).AlignmentStyle = ascVertical Then
                                ieTable = ieTable + String$(2, vbTab) + "<div style=""position:absolute" + _
                                    XYPos + _
                                    "; width:" & scWidth & "px" & _
                                    "; height:" & scHeight & "px" & _
                                    AddConditionalAttribute("; background-color", nBColor, False) + _
                                    ";"" id=""N" + CStr(sid) + """ " + IIf(UsesDynAPI, " h=0 v=0 sep=1", "") + "><div style=""position:absolute; top:" & Preferences.SepHeight \ 2 & "px; left:" & CInt(scWidth * ((100 - MenuCmds(c).SeparatorPercent) / 200)) & "px; width:" & CInt(scWidth * (MenuCmds(c).SeparatorPercent / 100)) & "px; height:1px; overflow:hidden; background-color:" & nTColor & """></div></div>"
                            Else
                                ieTable = ieTable + String$(2, vbTab) + "<div style=""position:absolute" + _
                                    XYPos + _
                                    "; width:" & scWidth & "px" & _
                                    "; height:" & scHeight + 2 * MenuGrps(g).CmdsFXSize & "px" & _
                                    AddConditionalAttribute("; background-color", nBColor, False) + _
                                    ";"" id=""N" + CStr(sid) + """ " + IIf(UsesDynAPI, " h=0 v=0 sep=1", "") + "><div style=""position:absolute; top:" & MenuGrps(g).CmdsFXSize + CInt(scHeight * ((100 - MenuCmds(c).SeparatorPercent) / 200)) & "px; left:" & Preferences.SepHeight \ 2 & "px; width:1px; height:" & CInt(scHeight * (MenuCmds(c).SeparatorPercent / 100)) & "px; overflow:hidden; background-color:" & nTColor & """></div></div>"
                            End If
                            'If UsesDynAPI Then ieTable = ieTable + "<div><div></div></div>"
                            sid = sid - 1
                        Else
                            ieTable = ieTable + String$(2, vbTab) + _
                                "<div style=""position:absolute" + _
                                XYPos + _
                                "; width:" & scWidth & "px" & _
                                "; height:" & scHeight + 2 * MenuCmds(c).CmdsFXSize & "px" & _
                                RaiusToCSS(MenuCmds(c).Radius) & _
                                "; white-space:nowrap" & _
                                IIf(LenB(MenuCmds(c).Caption) <> 0, _
                                    "; font-family:" + GetFontFamily(MenuCmds(c).NormalFont.FontName) + _
                                    "; font-size:" & MenuCmds(c).NormalFont.FontSize & "px" + _
                                    "; font-weight:" + IIf(MenuCmds(c).NormalFont.FontBold, "bold", "normal") + _
                                    "; font-style:" + IIf(MenuCmds(c).NormalFont.FontItalic, "italic", "normal") + _
                                    "; text-decoration:" + IIf(MenuCmds(c).NormalFont.FontUnderline, "underline", "none") + _
                                    "; color:" + nTColor, "; overflow:hidden;") + _
                                    CmdShadowToCSS(c, False) + cFX(1) + _
                                IIf(IsFlash(MenuCmds(c).BackImage.NormalImage), "", AddConditionalBackImg(MenuCmds(c).BackImage, False)) + _
                                AddConditionalAttribute("; background-color", nBColor, False) + _
                                IIf(MenuCmds(c).Disabled, "; cursor:" + cName, "") + _
                                ";"" id=""N" + CStr(id) + """" + _
                                IIf(UsesDynAPI, " h=" & MenuCmds(c).CmdsMarginX & " v=" & MenuCmds(c).CmdsMarginY, "") & _
                                IIf(MenuCmds(c).Disabled, ">", GetEventCode(c, False) + ">") + _
                                CreateImageTable(MenuCmds(c), False, False, scWidth, scHeight) + _
                            "</div>" + vbCrLf
                            
                            If Not MenuCmds(c).Disabled Then
                                ieTable = ieTable + String$(2, vbTab) + _
                                    "<div" + _
                                    IIf(dispTips And (LenB(MenuCmds(c).WinStatus) <> 0), " title=""" + ParseStatusText(c, False) + """", "") + _
                                    " style=""position:absolute" + _
                                    XYPos + _
                                    "; width:" & scWidth & "px" & _
                                    "; height:" & scHeight + 2 * MenuCmds(c).CmdsFXSize & "px" & _
                                    RaiusToCSS(MenuCmds(c).Radius) & _
                                    "; white-space:nowrap" & _
                                    IIf(LenB(MenuCmds(c).Caption) <> 0, _
                                        "; font-family:" + GetFontFamily(MenuCmds(c).HoverFont.FontName) + _
                                        "; font-size:" & MenuCmds(c).HoverFont.FontSize & "px" + _
                                        "; font-weight:" + IIf(MenuCmds(c).HoverFont.FontBold, "bold", "normal") + _
                                        "; font-style:" + IIf(MenuCmds(c).HoverFont.FontItalic, "italic", "normal") + _
                                        "; text-decoration:" + IIf(MenuCmds(c).HoverFont.FontUnderline, "underline", "none") + _
                                        "; color:" + hTColor, "; overflow:hidden;") + _
                                        CmdShadowToCSS(c, True) + cFX(2) + _
                                    IIf(IsFlash(MenuCmds(c).BackImage.HoverImage), "", AddConditionalBackImg(MenuCmds(c).BackImage, True)) + _
                                    AddConditionalAttribute("; background-color", hBColor, False) + _
                                    "; cursor:" + cName + _
                                    "; visibility:hidden;"" id=""O" + CStr(id) + """" + _
                                    IIf(MenuCmds(c).Disabled, ">", GetEventCode(c, True) + ">") + _
                                    CreateImageTable(MenuCmds(c), True, False, scWidth, scHeight) + _
                                "</div>" + vbCrLf
                            End If
                            
                            id = id + 1
                        End If
                    End If
                    
                    If Project.CompileNSCode Then
                        scWidth = scWidth - CalcFrameBorder(g, ascHorizontal)
                        If MenuGrps(g).AlignmentStyle = ascVertical Then
                            XYPos = " left=" & MenuGrps(g).FrameBorder + MenuGrps(g).ContentsMarginH - MenuGrps(g).FrameBorder & _
                                    " top=" & cmdTop
                        Else
                            XYPos = " left=" & cmdTop & _
                                    " top=" & MenuGrps(g).FrameBorder + MenuGrps(g).ContentsMarginV - MenuGrps(g).FrameBorder
                        End If
                        If MenuCmds(c).Name = "[SEP]" Then
                            If MenuGrps(g).AlignmentStyle = ascVertical Then
                                nsTable = nsTable + String$(1, vbTab) + "<layer" + _
                                    XYPos + _
                                    " width=" & scWidth & _
                                    " height=" & scHeight + 2 * MenuCmds(c).CmdsFXSize & _
                                    " z-index=1002" + _
                                    AddConditionalAttribute(" bgColor", nBColor, True) + ">" + _
                                "<layer left=" & CInt(scWidth * (100 - MenuCmds(c).SeparatorPercent) / 200) & " top=" & Preferences.SepHeight \ 2 & " width=" & CInt(scWidth * (MenuCmds(c).SeparatorPercent / 100)) & " height=1 bgcolor=" & nTColor & "></layer></layer>"
                            Else
                                nsTable = nsTable + String$(1, vbTab) + "<layer" + _
                                    XYPos + _
                                    " width=" & scWidth & _
                                    " height=" & scHeight + 2 * MenuCmds(c).CmdsFXSize & _
                                    " z-index=1002" + _
                                    AddConditionalAttribute(" bgColor", nBColor, True) + ">" + _
                                "<layer left=" & Preferences.SepHeight \ 2 & " top=" & CInt(scWidth * (100 - MenuCmds(c).SeparatorPercent) / 200) & " width=1 height=" & CInt(scHeight * (MenuCmds(c).SeparatorPercent / 100)) & " bgcolor=" & nTColor & "></layer></layer>"
                            End If
                        Else
                            nsEventCode = GetNSEventCode(c)
                            CmdLayerName = "MC" & lIdx
                            
                            nsImgTblNorm = CreateImageTable(MenuCmds(c), False, True, scWidth, scHeight)
                            nLayers(lIdx) = IIf(LenB(MenuCmds(c).Caption) <> 0, _
                            "<ilayer left=" & MenuCmds(c).CmdsMarginX & " top=" & ((scHeight + 2 * MenuCmds(c).CmdsFXSize) - (scHeight - 2 * MenuCmds(c).CmdsMarginY)) / 2 & ">" + _
                                IIf(MenuCmds(c).NormalFont.FontBold, "<b>", "") + _
                                IIf(MenuCmds(c).NormalFont.FontItalic, "<i>", "") + _
                                IIf(MenuCmds(c).NormalFont.FontUnderline, "<u>", "") + _
                                "<font face=" + GetFontFamily(MenuCmds(c).NormalFont.FontName, True) + _
                                    " point-size=" & px2pt(MenuCmds(c).NormalFont.FontSize) & _
                                    " color=" + nTColor + ">" + _
                                    nsImgTblNorm + _
                                "</font>" + _
                                IIf(MenuCmds(c).NormalFont.FontUnderline, "</u>", "") + _
                                IIf(MenuCmds(c).NormalFont.FontItalic, "</i>", "") + _
                                IIf(MenuCmds(c).NormalFont.FontBold, "</b>", ""), nsImgTblNorm) + _
                            IIf(LenB(MenuCmds(c).Caption) <> 0, "</ilayer>", "")
                            
                            If Not MenuCmds(c).Disabled Then
                                nsImgTblHovr = CreateImageTable(MenuCmds(c), True, True, scWidth, scHeight)
                                hLayers(lIdx) = IIf(LenB(MenuCmds(c).Caption) <> 0, _
                                "<ilayer left=" & MenuCmds(c).CmdsMarginX & " top=" & ((scHeight + 2 * MenuCmds(c).CmdsFXSize) - (scHeight - 2 * MenuCmds(c).CmdsMarginY)) / 2 & ">" + _
                                    IIf(MenuCmds(c).HoverFont.FontBold, "<b>", "") + _
                                    IIf(MenuCmds(c).HoverFont.FontItalic, "<i>", "") + _
                                    IIf(MenuCmds(c).HoverFont.FontUnderline, "<u>", "") + _
                                    "<font face=" + GetFontFamily(MenuCmds(c).HoverFont.FontName, True) + _
                                        " point-size=" & px2pt(MenuCmds(c).HoverFont.FontSize) & _
                                        " color=" + hTColor + ">" + _
                                            nsImgTblHovr + _
                                    "</font>" + _
                                    IIf(MenuCmds(c).HoverFont.FontUnderline, "</u>", "") + _
                                    IIf(MenuCmds(c).HoverFont.FontItalic, "</i>", "") + _
                                    IIf(MenuCmds(c).HoverFont.FontBold, "</b>", ""), nsImgTblHovr) + _
                                IIf(LenB(MenuCmds(c).Caption) <> 0, "</ilayer>", "")
                            End If
                            
                            scWidth = scWidth + CalcFrameBorder(g, ascHorizontal)
                            If Not MenuCmds(c).Disabled Then
                                nsTable = nsTable + String$(1, vbTab) + _
                                    "<layer name=" + CmdLayerName + "EH" + CStr(lIdx) + _
                                        XYPos & _
                                        " width=" & scWidth & _
                                        " height=" & scHeight + 2 * MenuCmds(c).CmdsFXSize & _
                                        " z-index=1003"
                                        If Not MenuCmds(c).Disabled Then
                                            nsTable = nsTable + nsEventCode(1)
                                        End If
                                        nsTable = nsTable + "></layer>"
                                        
                                If LenB(nsEventCode(2)) <> 0 Then
                                    nsEC = Mid(nsEventCode(2), 11)
                                    nsEC = Replace(Left(nsEC, Len(nsEC) - 1), """", "\""")
                                    nsStyleCode = nsStyleCode + String$(1, vbTab) + "nTCode[" + CStr(lIdx) + "]=""" + nsEC + """;" + vbCrLf
                                End If
                            End If
        
                            nsTable = nsTable + String$(1, vbTab) + _
                                "<layer name=" + CmdLayerName + "N" + _
                                    XYPos & _
                                    " width=" & scWidth & _
                                    " height=" & scHeight + 2 * MenuCmds(c).CmdsFXSize & _
                                    " z-index=1002" + _
                                    IIf(IsFlash(MenuCmds(c).BackImage.NormalImage), "", AddConditionalAttribute(" background", ImgAbsPath + GetFileName(MenuCmds(c).BackImage.NormalImage), True, True)) + _
                                    AddConditionalAttribute(" bgColor", nBColor, True) + _
                                    ">" + nLayers(lIdx) + "</layer>"
                            If Not MenuCmds(c).Disabled Then
                                nsTable = nsTable + String$(1, vbTab) + _
                                    "<layer name=" + CmdLayerName + "O" + _
                                        XYPos & _
                                        " width=" & scWidth & _
                                        " height=" & scHeight + 2 * MenuCmds(c).CmdsFXSize & _
                                        " z-index=1002" + _
                                        IIf(IsFlash(MenuCmds(c).BackImage.HoverImage), "", AddConditionalAttribute(" background", ImgAbsPath + GetFileName(MenuCmds(c).BackImage.HoverImage), True, True)) + _
                                        AddConditionalAttribute(" bgColor", hBColor, True) + _
                                        " visibility=hidden>" + hLayers(lIdx) + _
                                    "</layer>"
                            End If
                        End If
                    End If
                    
                    'bStyleCode = bStyleCode + ImageObj(MenuCmds(c))
                    
                    cmdTop = cmdTop + _
                             IIf(MenuGrps(g).AlignmentStyle = ascVertical, scHeight, scWidth) + _
                             MenuGrps(g).Leading + IIf(MenuCmds(c).Name = "[SEP]", 0, IIf(MenuGrps(g).AlignmentStyle = ascVertical, 2 * MenuCmds(c).CmdsFXSize, 0))
                End If
            Next c
            
            ieTable = ieTable + "</div></div>" + vbCrLf
            ieTable = ieTable + vbCrLf
            
            nsTable = nsTable + vbCrLf + "</layer></layer>" + vbCrLf
            nsTable = nsTable + vbCrLf
            
            SaveFile fnp + "iet" & g, ieTable
            SaveFile fnp + "nst" & g, nsTable
            ieTable = ""
            nsTable = ""
        End If
    Next g
    
    On Error Resume Next
    For g = 1 To UBound(MenuGrps)
        If MenuGrps(g).Compile Then
            ieTable = ieTable + LoadFile(fnp + "iet" & g)
            nsTable = nsTable + LoadFile(fnp + "nst" & g)
            Kill fnp + "iet" & g
            Kill fnp + "nst" & g
        End If
    Next g
    On Error GoTo 0
    Err.Clear
    
    ieStyleCode = Replace(ieStyleCode, ";;", ";")
    If Project.CodeOptimization <> cocDEBUG And Not IsForPreviewing Then
        'If Project.CompileIECode Or IsForPreviewing Then ieStyleCode = OptimizeStyleCode(ieStyleCode, 1)
        'If Project.CompileNSCode Or IsForPreviewing Then nsStyleCode = OptimizeStyleCode(nsStyleCode, 2)
        'bStyleCode = OptimizeImagesCode(bStyleCode)
    End If
    
    'Save the code...
    If IsForPreviewing Then
        SaveCode "menu", _
                        BuildCode(tpcBoth, _
                            IsForPreviewing, _
                            ieTable, nsTable, _
                            ieStyleCode, nsStyleCode, bStyleCode, True), hRefStatements, True, tpcBoth, False, GenDynAPICode(ieTable)
    Else
        Select Case Project.CodeOptimization
            Case cocDEBUG
                SaveCode Project.JSFileName, BuildCode(tpcBoth, IsForPreviewing, ieTable, nsTable, ieStyleCode, nsStyleCode, bStyleCode, Not FramesMode), hRefStatements, False, tpcBoth, DoNotSave, IIf(FramesMode, "", GenDynAPICode(ieTable))
                If FramesMode Then SaveCode Project.JSFileName + "_frames", scriptExtraFramesCode(IIf(Project.CompileIECode, GenBrowserSaveCode(ieTable, "IE", False), "") + IIf(Project.CompileNSCode, GenBrowserSaveCode(nsTable, "NS", False), ""), IsForPreviewing, tpcBoth), hRefStatements, False, tpcBoth, DoNotSave, GenDynAPICode(ieTable)
            Case Else
                If Project.CompileIECode Then
                    SetupExpandableKeywords tpcIE
                    SaveCode "ie" + Project.JSFileName, BuildCode(tpcIE, False, ieTable, nsTable, ieStyleCode, nsStyleCode, bStyleCode, Not FramesMode), hRefStatements, False, tpcIE, DoNotSave, IIf(FramesMode, "", OptimizeJSCode(GenDynAPICode(ieTable)))
                End If
                If Project.CompileNSCode Then
                    SetupExpandableKeywords tpcNS
                    SaveCode "ns" + Project.JSFileName, BuildCode(tpcNS, False, ieTable, nsTable, ieStyleCode, nsStyleCode, bStyleCode, Not FramesMode), hRefStatements, False, tpcNS, DoNotSave, ""
                End If
                If FramesMode Then
                    If Project.CompileIECode Then
                        SetupExpandableKeywords tpcIE
                        SaveCode "ie" + Project.JSFileName + "_frames", scriptExtraFramesCode(GenBrowserSaveCode(ieTable, "IE", False), False, tpcIE), hRefStatements, False, tpcIE, DoNotSave, OptimizeJSCode(GenDynAPICode(ieTable))
                    End If
                    If Project.CompileNSCode Then
                        SetupExpandableKeywords tpcNS
                        SaveCode "ns" + Project.JSFileName + "_frames", scriptExtraFramesCode(GenBrowserSaveCode(nsTable, "NS", False), False, tpcNS), hRefStatements, False, tpcNS, DoNotSave, ""
                    End If
                End If
        End Select
        PD = Project
    End If
             
ExitSub:
    
    FloodPanel.Value = 0
    If Not IsLivePreview Then frmMain.MousePointer = vbDefault
    mIsCompiling = False
    CompileProject = True

End Function

Private Function CmdShadowToCSS(cmdID As Integer, isOver As Boolean) As String

    Dim css As String

    If isOver Then
        With MenuCmds(cmdID).HoverFont.FontShadow
            If .Enabled1 Then css = css + ";text-shadow:" & GetRGB(.Color1) & " " & .OffsetX1 & "px " & .OffsetY1 & "px " & IIf(.Blur1 <> 0, .Blur1 & "px", "")
            If .Enabled2 Then
                If .Enabled1 Then
                    css = css & ","
                Else
                    css = css & ";text-shadow:"
                End If
                css = css & GetRGB(.Color2) & " " & .OffsetX2 & "px " & .OffsetY2 & "px " & IIf(.Blur2 <> 0, .Blur2 & "px", "")
            End If
        End With
    Else
        With MenuCmds(cmdID).NormalFont.FontShadow
            If .Enabled1 Then css = css + ";text-shadow:" & GetRGB(.Color1) & " " & .OffsetX1 & "px " & .OffsetY1 & "px " & IIf(.Blur1 <> 0, .Blur1 & "px", "")
            If .Enabled2 Then
                If .Enabled1 Then
                    css = css & ","
                Else
                    css = css & ";text-shadow:"
                End If
                css = css & GetRGB(.Color2) & " " & .OffsetX2 & "px " & .OffsetY2 & "px " & IIf(.Blur2 <> 0, .Blur2 & "px", "")
            End If
        End With
    End If
    
    If css <> "" Then css = css + ";"
    CmdShadowToCSS = css

End Function

Private Function GrpShadowToCSS(grpID As Integer, isOver As Boolean) As String

    Dim css As String

    If isOver Then
        With MenuGrps(grpID).DefHoverFont.FontShadow
            If .Enabled1 Then css = css + ";text-shadow:" & GetRGB(.Color1) & " " & .OffsetX1 & "px " & .OffsetY1 & "px " & IIf(.Blur1 <> 0, .Blur1 & "px", "")
            If .Enabled2 Then
                If .Enabled1 Then
                    css = css & ","
                Else
                    css = css & ";text-shadow:"
                End If
                css = css & GetRGB(.Color2) & " " & .OffsetX2 & "px " & .OffsetY2 & "px " & IIf(.Blur2 <> 0, .Blur2 & "px", "")
            End If
        End With
    Else
        With MenuGrps(grpID).DefNormalFont.FontShadow
            If .Enabled1 Then css = css + ";text-shadow:" & GetRGB(.Color1) & " " & .OffsetX1 & "px " & .OffsetY1 & "px " & IIf(.Blur1 <> 0, .Blur1 & "px", "")
            If .Enabled2 Then
                If .Enabled1 Then
                    css = css & ","
                Else
                    css = css & ";text-shadow:"
                End If
                css = css & GetRGB(.Color2) & " " & .OffsetX2 & "px " & .OffsetY2 & "px " & IIf(.Blur2 <> 0, .Blur2 & "px", "")
            End If
        End With
    End If
    
    If css <> "" Then css = css + ";"
    GrpShadowToCSS = css

End Function

Private Function IsMasterProject() As Boolean

    On Error GoTo IsNoMaster
    
    IsMasterProject = (UBound(Project.SecondaryProjects) > 0)
    
    Exit Function
    
IsNoMaster:
    IsMasterProject = False

End Function

Private Function MozillaAndSafariTrans(ByVal v As Integer) As String

    MozillaAndSafariTrans = "opacity:" + Replace(CStr((100 - v) / 100), ",", ".") + "; "

End Function

Private Function MozillaAndSafariShadow(ByVal v As Integer, c As Long) As String

    MozillaAndSafariShadow = "box-shadow:" & v & "px " & v & "px " & CInt(v / 4) & "px " & GetRGBA(c, 0.5) & "; "

End Function

Private Function AddConditionalBackImg(img As tImage, isOver As Boolean) As String

    Dim r As String
    
    If isOver Then
        If img.HoverImage <> "" Then
            r = AddConditionalAttribute("; background-image", _
                    "url(" + ImgAbsPath + GetFileName(img.HoverImage) + ")", False)
        End If
    Else
        If img.NormalImage <> "" Then
            r = AddConditionalAttribute("; background-image", _
                    "url(" + ImgAbsPath + GetFileName(img.NormalImage) + ")", False)
        End If
    End If
    If r <> "" Then
        If img.Tile = False Then r = r + "; background-repeat:no-repeat; background-position:center"
    End If

    AddConditionalBackImg = r

End Function

Private Function AddConditionalAttribute(AttribName As String, ByVal Value As String, IsNS As Boolean, Optional AddDoubleQuotes As Boolean) As String

    On Error Resume Next
    Dim p1 As Integer
    Dim p2 As Integer

    If LenB(Value) <> 0 And Value <> "url()" And Value <> "url(%%REL%%)" And Value <> "%%REL%%" And Value <> ImgAbsPath Then
        If InStr(LCase(AttribName), "color") > 0 And Val(Value) = -2 Then Exit Function
        
        p1 = InStrRev(Value, "/")
        p2 = InStrRev(Value, "\")
        If p2 > p1 Then p1 = p2
        If p1 <> 0 Then If InStr(p1, Value, ".", vbTextCompare) = 0 Then Exit Function
        
        AddDoubleQuotes = AddDoubleQuotes Or (IsNS And (InStr(Value, "/") Or InStr(Value, "\")))
        AddConditionalAttribute = AttribName + IIf(IsNS, "=" + IIf(AddDoubleQuotes, """", ""), ":") + Value + IIf(AddDoubleQuotes, """", "")
    End If

End Function

Private Function ParseStatusText(idx As Integer, IsGroup As Boolean, Optional IsNS As Boolean) As String

    Dim s As String
    Dim c As String
    Dim u As String
    
    If IsGroup Then
        With MenuGrps(idx)
            s = .WinStatus
            c = .Caption
            If .Actions.OnClick.Type = atcNewWindow Or .Actions.OnClick.Type = atcURL Then
                u = .Actions.OnClick.url
            End If
        End With
    Else
        With MenuCmds(idx)
            s = .WinStatus
            c = .Caption
            If .Actions.OnClick.Type = atcNewWindow Or .Actions.OnClick.Type = atcURL Then
                u = .Actions.OnClick.url
            End If
        End With
    End If
    
    s = Replace(s, "%c", ParseCaptionHTML(c))
    s = Replace(s, "%u", FixURL(u))
    
    ParseStatusText = EncodeStr(PrepareCaption(s, IsNS))

End Function

Private Function GetFontFamily(FontName As String, Optional IsNS As Boolean) As String

    Dim fs() As String
    Dim i As Integer
    
    On Error GoTo ForceAbort
        
    If InStr(Project.FontSubstitutions, "|") > 0 Then
        fs = Split(Project.FontSubstitutions, "|")
        For i = 0 To UBound(fs) \ 2
            If fs(i * 2) = FontName Then
                GetFontFamily = IIf(InStr(FontName, " "), "'" + FontName + "'", FontName) + ", " + fs(i * 2 + 1)
                Exit For
            End If
        Next i
    Else
        GetFontFamily = IIf(InStr(FontName, " "), "'" + FontName + "'", FontName)
    End If
    
CleanExit:
    If LenB(GetFontFamily) = 0 Then GetFontFamily = FontName
    
    If IsNS Then
        GetFontFamily = Replace(IIf(InStr(GetFontFamily, " "), """" + GetFontFamily + """", GetFontFamily), "'", "")
    End If
    
    Exit Function
    
ForceAbort:
    GetFontFamily = ""
    GoTo CleanExit
    
End Function

Private Function GetBorderStyleName(ByVal idx As CommandsFXConstants) As String

    Select Case idx
        Case cfxcRaised: GetBorderStyleName = "outset"
        Case cfxcSunken: GetBorderStyleName = "groove"
        Case cfxcDouble: GetBorderStyleName = "double"
        Case cfxcBevel: GetBorderStyleName = "ridge"
        Case cfxcDotted: GetBorderStyleName = "dotted"
        Case cfxcDashed: GetBorderStyleName = "dashed"
        Case Else: GetBorderStyleName = "solid"
    End Select

End Function

Private Function GetBorderFX(i As Integer, IsGroup As Boolean, Optional IsScroll As Boolean = False, Optional IsGlobal As Boolean = False) As String()

    Dim FXColor As String
    Dim cFX(1 To 2) As String
    Dim g As MenuGrp
    Dim c As MenuCmd
    
    If IsGlobal Then
        g.CmdsFXhColor = Project.AutoScroll.FXhColor
        g.CmdsFXnColor = Project.AutoScroll.FXnColor
        g.CmdsFXNormal = Project.AutoScroll.FXNormal
        g.CmdsFXOver = Project.AutoScroll.FXOver
        g.CmdsFXSize = Project.AutoScroll.FXSize
    Else
        If IsGroup Then
            g = MenuGrps(i)
            If IsScroll Then
                g.CmdsFXhColor = g.Scrolling.FXhColor
                g.CmdsFXnColor = g.Scrolling.FXnColor
                g.CmdsFXNormal = g.Scrolling.FXNormal
                g.CmdsFXOver = g.Scrolling.FXOver
                g.CmdsFXSize = g.Scrolling.FXSize
            End If
        Else
            c = MenuCmds(i)
        End If
    End If
    
    If IsGroup Then
        With g
            If .CmdsFXSize > 0 Then
                FXColor = GetRGB(.CmdsFXnColor, True)
                Select Case .CmdsFXNormal
                    Case cfxcNone
                        If .CmdsFXnColor <> -2 Then cFX(1) = "border:" & .CmdsFXSize & "px solid " + FXColor
                    Case cfxcRaised
                        If .CmdsFXnColor = -2 Then
                            cFX(1) = "border:" & .CmdsFXSize & "px outset " + FXColor
                        Else
                            cFX(1) = GetBorderHTML(.CmdsFXSize, LightenColor(.CmdsFXnColor), LightenColor(.CmdsFXnColor), DarkenColor(.CmdsFXnColor), DarkenColor(.CmdsFXnColor))
                        End If
                    Case cfxcSunken
                        If .CmdsFXnColor = -2 Then
                            cFX(1) = "border:" & .CmdsFXSize & "px groove " + FXColor
                        Else
                            cFX(1) = GetBorderHTML(.CmdsFXSize, DarkenColor(.CmdsFXnColor), DarkenColor(.CmdsFXnColor), LightenColor(.CmdsFXnColor), LightenColor(.CmdsFXnColor))
                        End If
                    Case cfxcDouble
                        cFX(1) = "border:" & .CmdsFXSize & "px double " + FXColor
                    Case cfxcBevel
                        cFX(1) = "border:" & .CmdsFXSize & "px ridge " + FXColor
                    Case cfxcDotted
                        cFX(1) = "border:" & .CmdsFXSize & "px dotted " + FXColor
                    Case cfxcDashed
                        cFX(1) = "border:" & .CmdsFXSize & "px dashed " + FXColor
                End Select
                
                cFX(2) = ""
                FXColor = GetRGB(.CmdsFXhColor, True)
                Select Case .CmdsFXOver
                    Case cfxcNone
                        If .CmdsFXhColor <> -2 Then cFX(2) = "border:" & .CmdsFXSize & "px solid " + FXColor
                    Case cfxcRaised
                        If .CmdsFXhColor = -2 Then
                            cFX(2) = "border:" & .CmdsFXSize & "px outset " + FXColor
                        Else
                            cFX(2) = GetBorderHTML(.CmdsFXSize, LightenColor(.CmdsFXhColor), LightenColor(.CmdsFXhColor), DarkenColor(.CmdsFXhColor), DarkenColor(.CmdsFXhColor))
                        End If
                    Case cfxcSunken
                        If .CmdsFXhColor = -2 Then
                            cFX(2) = "border:" & .CmdsFXSize & "px groove " + FXColor
                        Else
                            cFX(2) = GetBorderHTML(.CmdsFXSize, DarkenColor(.CmdsFXhColor), DarkenColor(.CmdsFXhColor), LightenColor(.CmdsFXhColor), LightenColor(.CmdsFXhColor))
                        End If
                    Case cfxcDouble
                        cFX(2) = "border:" & .CmdsFXSize & "px double " + FXColor
                    Case cfxcBevel
                        cFX(2) = "border:" & .CmdsFXSize & "px ridge " + FXColor
                    Case cfxcDotted
                        cFX(2) = "border:" & .CmdsFXSize & "px dotted " + FXColor
                    Case cfxcDashed
                        cFX(2) = "border:" & .CmdsFXSize & "px dashed " + FXColor
                End Select
            End If
        End With
    Else
        With c
            FXColor = GetRGB(.CmdsFXnColor, True)
            Select Case .CmdsFXNormal
                Case cfxcNone
                    If .CmdsFXnColor <> -2 Then cFX(1) = "border:" & .CmdsFXSize & "px solid " + FXColor
                Case cfxcRaised
                    If .CmdsFXnColor = -2 Then
                        cFX(1) = "border:" & .CmdsFXSize & "px outset " + FXColor
                    Else
                        cFX(1) = GetBorderHTML(.CmdsFXSize, LightenColor(.CmdsFXnColor), LightenColor(.CmdsFXnColor), DarkenColor(.CmdsFXnColor), DarkenColor(.CmdsFXnColor))
                    End If
                Case cfxcSunken
                    If .CmdsFXnColor = -2 Then
                        cFX(1) = "border:" & .CmdsFXSize & "px groove " + FXColor
                    Else
                        cFX(1) = GetBorderHTML(.CmdsFXSize, DarkenColor(.CmdsFXnColor), DarkenColor(.CmdsFXnColor), LightenColor(.CmdsFXnColor), LightenColor(.CmdsFXnColor))
                    End If
                Case cfxcDouble
                    cFX(1) = "border:" & .CmdsFXSize & "px double " + FXColor
                Case cfxcBevel
                    cFX(1) = "border:" & .CmdsFXSize & "px ridge " + FXColor
                Case cfxcDotted
                    cFX(1) = "border:" & .CmdsFXSize & "px dotted " + FXColor
                Case cfxcDashed
                    cFX(1) = "border:" & .CmdsFXSize & "px dashed " + FXColor
            End Select
            
            cFX(2) = ""
            FXColor = GetRGB(.CmdsFXhColor, True)
            Select Case .CmdsFXOver
                Case cfxcNone
                    If .CmdsFXhColor <> -2 Then cFX(2) = "border:" & .CmdsFXSize & "px solid " + FXColor
                Case cfxcRaised
                    If .CmdsFXhColor = -2 Then
                        cFX(2) = "border:" & .CmdsFXSize & "px outset " + FXColor
                    Else
                        cFX(2) = GetBorderHTML(.CmdsFXSize, LightenColor(.CmdsFXhColor), LightenColor(.CmdsFXhColor), DarkenColor(.CmdsFXhColor), DarkenColor(.CmdsFXhColor))
                    End If
                Case cfxcSunken
                    If .CmdsFXhColor = -2 Then
                        cFX(2) = "border:" & .CmdsFXSize & "px groove " + FXColor
                    Else
                        cFX(2) = GetBorderHTML(.CmdsFXSize, DarkenColor(.CmdsFXhColor), DarkenColor(.CmdsFXhColor), LightenColor(.CmdsFXhColor), LightenColor(.CmdsFXhColor))
                    End If
                Case cfxcDouble
                    cFX(2) = "border:" & .CmdsFXSize & "px double " + FXColor
                Case cfxcBevel
                    cFX(2) = "border:" & .CmdsFXSize & "px ridge " + FXColor
                Case cfxcDotted
                    cFX(2) = "border:" & .CmdsFXSize & "px dotted " + FXColor
                Case cfxcDashed
                    cFX(2) = "border:" & .CmdsFXSize & "px dashed " + FXColor
            End Select
        End With
    End If
    
    If cFX(1) <> "" Then cFX(1) = "; " + cFX(1)
    If cFX(2) <> "" Then cFX(2) = "; " + cFX(2)
    
    GetBorderFX = cFX
    
End Function

Private Function GetBorderHTML(s As Integer, l As Long, t As Long, r As Long, b As Long) As String

    GetBorderHTML = "border-left:" & s & "px solid " & GetRGB(l, True) & _
                    ";border-top:" & s & "px solid " & GetRGB(t, True) & _
                    ";border-right:" & s & "px solid " & GetRGB(r, True) & _
                    ";border-bottom:" & s & "px solid " & GetRGB(b, True)

End Function

'Private Function scriptSounds(BrowserVar As String, UseExpand As Boolean) As String
'
'    Dim i As Integer
'    Dim sStr As String
'    Dim sCodeIE As String
'    Dim sCodeNS As String
'
'    For i = 1 To UBound(Sounds)
'        If BrowserVar = "IE" Then
'            sCodeIE = sCodeIE + "<embed src=""" + ImgAbsPath + GetFileName(Sounds(i)) + """ type=""video/x-ms-asf"" autostart=false loop=false volume=100 hidden name=Snd" & i & ">"
'        End If
'        If BrowserVar = "NS" Then
'            sCodeNS = sCodeNS + "<embed src=""" + ImgAbsPath + GetFileName(Sounds(i)) + """ autostart=false loop=false volume=100 hidden MASTERSOUND name=Snd" & i & ">"
'        End If
'    Next i
'
'    If BrowserVar = "IE" Then
'        sStr = sStr + "if(IsWin&&IE&&!SM)" + vbCrLf
'        sStr = sStr + "write(" + IIf(UseExpand, "Expand(""", """") + EncodeStr("<div style=""position:absolute;"">" + sCodeIE + "</div>") + IIf(UseExpand, """)", """") + ");"
'    End If
'
'    If BrowserVar = "NS" Then
'        sStr = sStr + "if(IsWin&&NS)" + vbCrLf
'        'sStr = sStr + "write(" + IIf(UseExpand, "Expand(""", """") + EncodeStr("<layer>" + sCodeNS + "</layer>") + IIf(UseExpand, """)", """") + ");"
'        sStr = sStr + "write(" + IIf(UseExpand, "Expand(""", """") + EncodeStr(sCodeNS) + IIf(UseExpand, """)", """") + ");"
'    End If
'
'    scriptSounds = sStr
'
'End Function

'Private Function GetSoundID(SoundFileName As String) As Integer
'
'    Dim i As Integer
'
'    For i = 1 To UBound(Sounds)
'        If Sounds(i) = SoundFileName Then
'            GetSoundID = i
'            Exit For
'        End If
'    Next i
'
'End Function

'Private Sub Add2SoundsArray(SoundFileName As String)
'
'    Dim i As Integer
'
'    If SoundFileName <> "" Then
'
'        For i = 1 To UBound(Sounds)
'            If Sounds(i) = SoundFileName Then Exit Sub
'        Next i
'
'        ReDim Preserve Sounds(UBound(Sounds) + 1)
'        Sounds(UBound(Sounds)) = SoundFileName
'
'    End If
'
'End Sub

Public Function GetCursorName(dCursor As tCursor) As String

    Select Case dCursor.cType
        Case iccDefault
            GetCursorName = "default"
        Case iccCrosshair
            GetCursorName = "crosshair"
        Case iccHand
            GetCursorName = "pointer"
        Case iccText
            GetCursorName = "text"
        Case iccHelp
            GetCursorName = "help"
        Case iccResizeE
            GetCursorName = "e-resize"
        Case iccResizeNE
            GetCursorName = "ne-resize"
        Case iccResizeNW
            GetCursorName = "nw-resize"
        Case iccResizeN
            GetCursorName = "n-resize"
        Case iccResizeSE
            GetCursorName = "se-resize"
        Case iccResizeSW
            GetCursorName = "sw-resize"
        Case iccResizeS
            GetCursorName = "s-resize"
        Case iccResizeW
            GetCursorName = "w-resize"
        Case iccResizeAll
            GetCursorName = "move"
        Case iccWait
            GetCursorName = "wait"
        Case iccCustom
            GetCursorName = "url(" + ImgAbsPath + GetFileName(dCursor.cFile) + ")"
    End Select

End Function

Private Function RaiusToCSS(Radius As BorderRadius) As String

    Dim sStr As String
    
    sStr = "; border-radius: %1px %2px %3px %4px"
    sStr = Replace(sStr, "%1", Radius.TopLeft)
    sStr = Replace(sStr, "%2", Radius.TopRight)
    sStr = Replace(sStr, "%3", Radius.BottomRight)
    sStr = Replace(sStr, "%4", Radius.BottomLeft)
    
    If sStr = "; border-radius: 0px 0px 0px 0px" Then sStr = ""
    
    RaiusToCSS = sStr

End Function

Private Function CreateCorners(g As Integer, DivWidth As Integer, DivHeight As Integer, IsNS As Boolean) As String

    Dim table As String
    
    'Top Left
    Dim tll As Integer
    Dim tlt As Integer
    Dim tlw As Integer
    Dim tlh As Integer
    
    'Top Center
    Dim tcl As Integer
    Dim tct As Integer
    Dim tcw As Integer
    Dim tch As Integer
    
    'Top Right
    Dim trl As Integer
    Dim trt As Integer
    Dim trw As Integer
    Dim trh As Integer
    
    'Center Left
    Dim cll As Integer
    Dim clt As Integer
    Dim clw As Integer
    Dim clh As Integer
    
    'Center Right
    Dim crl As Integer
    Dim crt As Integer
    Dim crw As Integer
    Dim crh As Integer
    
    'Bottom Left
    Dim bll As Integer
    Dim blt As Integer
    Dim blw As Integer
    Dim blh As Integer
    
    'Bottom Center
    Dim bcl As Integer
    Dim bct As Integer
    Dim bcw As Integer
    Dim bch As Integer
    
    'Bottom Right
    Dim brl As Integer
    Dim brt As Integer
    Dim brw As Integer
    Dim brh As Integer
    
    With MenuGrps(g).CornersImages
    
        'Center Width/Height
        clw = GetImageSize(.gcLeft)(0)
        clh = DivHeight
        crw = GetImageSize(.gcRight)(0)
        crh = DivHeight
    
        'Top Width/Height
        tlw = GetImageSize(.gcTopLeft)(0)
        tlh = GetImageSize(.gcTopLeft)(1)
        trw = GetImageSize(.gcTopRight)(0)
        trh = GetImageSize(.gcTopRight)(1)
        tcw = DivWidth + IIf(tlw = 0, leftCorner, 0) + IIf(trw = 0, rightCorner, 0)
        tch = GetImageSize(.gcTopCenter)(1)
        
        'Bottom Width/Height
        blw = GetImageSize(.gcBottomLeft)(0)
        blh = GetImageSize(.gcBottomLeft)(1)
        brw = GetImageSize(.gcBottomRight)(0)
        brh = GetImageSize(.gcBottomRight)(1)
        bcw = DivWidth + IIf(blw = 0, leftCorner, 0) + IIf(brw = 0, rightCorner, 0)
        bch = GetImageSize(.gcBottomCenter)(1)
        
        'Center Adjustment
        clh = clh + IIf(tlh = 0, topCorner, 0) + IIf(blh = 0, bottomCorner, 0)
        crh = crh + IIf(trh = 0, topCorner, 0) + IIf(brh = 0, bottomCorner, 0)
        
        'Top Left/Top
        tll = leftCorner - tlw
        tlt = topCorner - tlh
        tcl = tll + tlw
        tct = topCorner - tch
        trl = leftCorner + DivWidth
        trt = topCorner - trh
        
        'Center Left/Top
        cll = leftCorner - clw
        clt = tlt + tlh
        crl = leftCorner + DivWidth
        crt = trt + trh
        
        'Bottom Left/Top
        bll = leftCorner - blw
        blt = topCorner + DivHeight
        bcl = bll + blw
        bct = topCorner + DivHeight
        brl = leftCorner + DivWidth
        brt = topCorner + DivHeight
        
        If IsNS Then
            'TOP CORNERs
            If LenB(.gcTopLeft) <> 0 Then
                table = table + "<layer" + _
                                " left=" & tll & _
                                " top=" & tlt & _
                                " width=" & tlw & _
                                " height=" & tlh & _
                                " background=""" + ImgAbsPath + GetFileName(.gcTopLeft) + """></layer>"
            End If
            If LenB(.gcTopCenter) <> 0 Then
                table = table + "<layer" + _
                                " left=" & tcl & _
                                " top=" & tct & _
                                " width=" & tcw & _
                                " height=" & tch & _
                                " background=""" + ImgAbsPath + GetFileName(.gcTopCenter) + """></layer>"
            End If
            If LenB(.gcTopRight) <> 0 Then
                table = table + "<layer" + _
                                " left=" & trl & _
                                " top=" & trt & _
                                " width=" & trw & _
                                " height=" & trh & _
                                " background=""" + ImgAbsPath + GetFileName(.gcTopRight) + """></layer>"
            End If
            
            'CENTER CORNERs
            If LenB(.gcLeft) <> 0 Then
                table = table + "<layer" + _
                                " left=" & cll & _
                                " top=" & clt & _
                                " width=" & clw & _
                                " height=" & clh & _
                                " background=""" + ImgAbsPath + GetFileName(.gcLeft) + """></layer>"
            End If
            If LenB(.gcRight) <> 0 Then
                table = table + "<layer" + _
                                " left=" & crl & _
                                " top=" & crt & _
                                " width=" & crw & _
                                " height=" & crh & _
                                " background=""" + ImgAbsPath + GetFileName(.gcRight) + """></layer>"
            End If
            
            'BOTTOM CORNERs
            If LenB(.gcBottomLeft) <> 0 Then
                table = table + "<layer" + _
                                " left=" & bll & _
                                " top=" & blt & _
                                " width=" & blw & _
                                " height=" & blh & _
                                " background=""" + ImgAbsPath + GetFileName(.gcBottomLeft) + """></layer>"
            End If
            If LenB(.gcBottomCenter) <> 0 Then
                table = table + "<layer" + _
                                " left=" & bcl & _
                                " top=" & bct & _
                                " width=" & bcw & _
                                " height=" & bch & _
                                " background=""" + ImgAbsPath + GetFileName(.gcBottomCenter) + """></layer>"
            End If
            If LenB(.gcBottomRight) <> 0 Then
                table = table + "<layer" + _
                                " left=" & brl & _
                                " top=" & brt & _
                                " width=" & brw & _
                                " height=" & brh & _
                                " background=""" + ImgAbsPath + GetFileName(.gcBottomRight) + """></layer>"
            End If
        Else
            'TOP CORNERs
            If LenB(.gcTopLeft) <> 0 Then
                table = table + "<span style=""position:absolute" + _
                                "; top:" & tlt & "px" + _
                                "; left:" & tll & "px" + _
                                "; width:" & tlw & "px" & _
                                "; height:" & tlh & "px" + _
                                "; background-image:url(" + ImgAbsPath + GetFileName(.gcTopLeft) + "); overflow:hidden; z-index:1""></span>"
            End If
            If LenB(.gcTopCenter) <> 0 Then
                table = table + "<span style=""position:absolute" + _
                                "; top:" & tct & "px" + _
                                "; left:" & tcl & "px" + _
                                "; width:" & tcw & "px" & _
                                "; height:" & tch & "px" + _
                                "; background-image:url(" + ImgAbsPath + GetFileName(.gcTopCenter) + "); overflow:hidden; z-index:1""></span>"
            End If
            If LenB(.gcTopRight) <> 0 Then
                table = table + "<span style=""position:absolute" + _
                                "; top:" & trt & "px" + _
                                "; left:" & trl & "px" + _
                                "; width:" & trw & "px" & _
                                "; height:" & trh & "px" + _
                                "; background-image:url(" + ImgAbsPath + GetFileName(.gcTopRight) + "); overflow:hidden; z-index:1""></span>"
            End If
            
            'CENTER CORNERs
            If LenB(.gcLeft) <> 0 Then
                table = table + "<span style=""position:absolute" + _
                                "; top:" & clt & "px" + _
                                "; left:" & cll & "px" + _
                                "; width:" & clw & "px" & _
                                "; height:" & clh & "px" + _
                                "; background-image:url(" + ImgAbsPath + GetFileName(.gcLeft) + "); overflow:hidden; z-index:1""></span>"
            End If
            If LenB(.gcRight) <> 0 Then
                table = table + "<span style=""position:absolute" + _
                                "; top:" & crt & "px" + _
                                "; left:" & crl & "px" + _
                                "; width:" & crw & "px" & _
                                "; height:" & crh & "px" + _
                                "; background-image:url(" + ImgAbsPath + GetFileName(.gcRight) + "); overflow:hidden; z-index:1""></span>"
            End If
            
            'BOTTOM CORNERs
            If LenB(.gcBottomLeft) <> 0 Then
                table = table + "<span style=""position:absolute; z-index:1" + _
                                "; top:" & blt & "px" + _
                                "; left:" & bll & "px" + _
                                "; width:" & blw & "px" & _
                                "; height:" & blh & "px" + _
                                "; background-image:url(" + ImgAbsPath + GetFileName(.gcBottomLeft) + "); overflow:hidden; z-index:1""></span>"
            End If
            If LenB(.gcBottomCenter) <> 0 Then
                table = table + "<span style=""position:absolute" + _
                                "; top:" & bct & "px" + _
                                "; left:" & bcl & "px" + _
                                "; width:" & bcw & "px" & _
                                "; height:" & bch & "px" + _
                                "; background-image:url(" + ImgAbsPath + GetFileName(.gcBottomCenter) + "); overflow:hidden; z-index:1""></span>"
            End If
            If LenB(.gcBottomRight) <> 0 Then
                table = table + "<span style=""position:absolute" + _
                                "; top:" & brt & "px" + _
                                "; left:" & brl & "px" + _
                                "; width:" & brw & "px" & _
                                "; height:" & brh & "px" + _
                                "; background-image:url(" + ImgAbsPath + GetFileName(.gcBottomRight) + "); overflow:hidden; z-index:1""></span>"
            End If
        End If
    End With
    
    CreateCorners = table

End Function

Private Function CreateImageTable(mc As MenuCmd, UseHoverImage As Boolean, IsNS As Boolean, ByVal DivWidth As Integer, ByVal CommandHeight As Integer) As String

    Dim table As String
    Dim imgFile As String
    Dim c As Integer
    Dim imgLHS As Integer
    Dim imgRHS As Integer
    Dim mg As MenuGrp
    Dim leftP As Integer
    Dim topP As Integer
    
    Dim lImage As String
    Dim rImage As String
    Dim txt As String
    
    Dim lp As Integer
    Dim rp As Integer
    
    Dim IsBorderRendered As Boolean
    
    c = GetIDByName(mc.Name)
    mg = MenuGrps(mc.Parent)
    imgLHS = IIf(LenB(mc.LeftImage.NormalImage) <> 0, Preferences.ImgSpace, 0)
    imgRHS = IIf(LenB(mc.RightImage.NormalImage) <> 0, Preferences.ImgSpace, 0)
    
    If IsNS Then
        leftP = mc.CmdsFXSize
        topP = 0
        
        DivWidth = DivWidth - 2 * leftP - 2 * mc.CmdsMarginX + CalcFrameBorder(mc.Parent, ascHorizontal)
        CommandHeight = CommandHeight - 2 * mc.CmdsMarginY

        table = "<layer" + _
                    " left=" & leftP & _
                    " top=" & topP & _
                    " width=" & DivWidth & _
                    " height=" & CommandHeight & _
                ">"
        With mc
            If UseHoverImage Then
                If IsFlash(.BackImage.HoverImage) Then
                    table = table + CreateFlashCode(.BackImage.HoverImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.hBackColor, True), True)
                End If
            Else
                If IsFlash(.BackImage.NormalImage) Then
                    table = table + CreateFlashCode(.BackImage.NormalImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.nBackColor, True), True)
                End If
            End If
            
            If LenB(.LeftImage.NormalImage) <> 0 Then
                If UseHoverImage Then
                    imgFile = ImgAbsPath + GetFileName(.LeftImage.HoverImage)
                Else
                    imgFile = ImgAbsPath + GetFileName(.LeftImage.NormalImage)
                End If
                lImage = "<layer top=" & CInt((CommandHeight - .LeftImage.h) / 2) & " left=0>"
                lImage = lImage + "<img src=""" + imgFile + """ width=" & .LeftImage.w & " height=" & .LeftImage.h & ">"
                lImage = lImage + "</layer>"
            End If
            
            If LenB(.Caption) <> 0 Then
                txt = "<layer top=" & CInt((CommandHeight - GetCmdTextHeight(c)) / 2)
                lp = .LeftImage.w + imgLHS
                rp = .RightImage.w + imgRHS
                txt = txt + " width=" & DivWidth - (lp + rp) & " left=" & lp & ">"
                Select Case .Alignment
                    Case tacLeft
                        txt = txt + "<div align=""left"">"
                    Case tacCenter
                        txt = txt + "<div align=""center"">"
                    Case tacRight
                        txt = txt + "<div align=""right"">"
                End Select
                txt = txt + PrepareCaption(.Caption, True) + "</div></layer>"
            End If
            
            If LenB(.RightImage.NormalImage) <> 0 Then
                If UseHoverImage Then
                    imgFile = ImgAbsPath + GetFileName(.RightImage.HoverImage)
                Else
                    imgFile = ImgAbsPath + GetFileName(.RightImage.NormalImage)
                End If
                rImage = "<layer top=" & CInt((CommandHeight - .RightImage.h) / 2) & " left=" & DivWidth - .RightImage.w & ">"
                rImage = rImage + "<img src=""" + imgFile + """ width=" & .RightImage.w & " height=" & .RightImage.h & ">"
                rImage = rImage + "</layer>"
            End If
            
            If LenB(lImage) <> 0 Or LenB(txt) <> 0 Or LenB(rImage) <> 0 Then
                table = table + lImage + txt + rImage + "</layer>"
            Else
                table = ""
            End If
        End With
    Else
        If UseHoverImage Then
            IsBorderRendered = (mc.CmdsFXhColor <> -2) Or (mc.CmdsFXOver <> cfxcNone)
            Select Case mc.CmdsFXOver
                Case cfxcNone
                    leftP = IIf(IsBorderRendered, 0, mc.CmdsFXSize) + mc.CmdsMarginX
                    topP = IIf(IsBorderRendered, 0, mc.CmdsFXSize) + mc.CmdsMarginY
                Case Else
                    leftP = mc.CmdsMarginX
                    topP = mc.CmdsMarginY
            End Select
        Else
            IsBorderRendered = (mc.CmdsFXnColor <> -2) Or (mc.CmdsFXNormal <> cfxcNone)
            Select Case mc.CmdsFXNormal
                Case cfxcNone
                    leftP = IIf(IsBorderRendered, 0, mc.CmdsFXSize) + mc.CmdsMarginX
                    topP = IIf(IsBorderRendered, 0, mc.CmdsFXSize) + mc.CmdsMarginY
                Case Else
                    leftP = mc.CmdsMarginX
                    topP = mc.CmdsMarginY
            End Select
        End If
        If Not IsBorderRendered Then
            If Int(mc.CmdsFXSize / 2) <> mc.CmdsFXSize / 2 Then
                topP = topP - 1
            End If
        End If
        
        DivWidth = DivWidth - 2 * leftP - IIf(IsBorderRendered, 2 * mc.CmdsFXSize, 0)
        CommandHeight = CommandHeight - 2 * topP
        
        With mc
            If UseHoverImage Then
                If IsFlash(.BackImage.HoverImage) Then
                    table = table + CreateFlashCode(.BackImage.HoverImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.hBackColor, True), False)
                End If
            Else
                If IsFlash(.BackImage.NormalImage) Then
                    table = table + CreateFlashCode(.BackImage.NormalImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.nBackColor, True), False)
                End If
            End If
        
            If LenB(.LeftImage.NormalImage) <> 0 Then
                imgFile = ImgAbsPath + GetFileName(IIf(UseHoverImage, .LeftImage.HoverImage, .LeftImage.NormalImage))
                lImage = lImage + "<img" + _
                                    IIf(UsesDynAPI, " id=""L"" ", " ") + _
                                    "style=""position:absolute; top:" & CInt((CommandHeight - .LeftImage.h) / 2) + topP + IIf(IsBorderRendered, 0, mc.CmdsFXSize) & "px; left:" & leftP + .LeftImage.Margin & "px;"" src=""" + imgFile + """ width=" & .LeftImage.w & " height=" & .LeftImage.h & ">"
            End If
            
            If LenB(.Caption) <> 0 Then
                txt = "<div style=""position:absolute; top:" & CInt((CommandHeight - GetCmdTextHeight(c)) / 2) + topP + IIf(IsBorderRendered, 0, mc.CmdsFXSize) & "px"
                lp = .LeftImage.w + imgLHS
                rp = .RightImage.w + imgRHS
                txt = txt + "; width:" & DivWidth - (lp + rp) & "px; left:" & lp + leftP & "px;"""
                Select Case .Alignment
                    Case tacLeft
                        txt = txt + " align=""left"""
                    Case tacCenter
                        txt = txt + " align=""center"""
                    Case tacRight
                        txt = txt + " align=""right"""
                End Select
                If UseHoverImage Then
                    txt = txt + "></div>"
                Else
                    txt = txt + ">" + PrepareCaption(.Caption, False) + "</div>"
                End If
            End If
            
            If LenB(.RightImage.NormalImage) <> 0 Then
                imgFile = ImgAbsPath + GetFileName(IIf(UseHoverImage, .RightImage.HoverImage, .RightImage.NormalImage))
                rImage = rImage + "<img" + _
                                    IIf(UsesDynAPI, " id=""R"" ", " ") + _
                                    "style=""position:absolute; top:" & CInt((CommandHeight - .RightImage.h) / 2) + topP + IIf(IsBorderRendered, 0, mc.CmdsFXSize) & "px; left:" & DivWidth - .RightImage.w + leftP - .RightImage.Margin & "px;"" src=""" + imgFile + """ width=" & .RightImage.w & " height=" & .RightImage.h & ">"
            End If
            
            table = table + lImage + txt + rImage '+ "</span>"
        End With
    End If
    
    CreateImageTable = table

End Function

'Private Function CreateImageTable(mc As MenuCmd, UseHoverImage As Boolean, IsNS As Boolean, ByVal DivWidth As Integer, ByVal CommandHeight As Integer) As String
'
'    Dim table As String
'    Dim imgFile As String
'    Dim c As Integer
'    Dim imgLHS As Integer
'    Dim imgRHS As Integer
'    Dim mg As MenuGrp
'    Dim leftP As Integer
'    Dim topP As Integer
'
'    Dim lImage As String
'    Dim rImage As String
'    Dim txt As String
'
'    Dim lp As Integer
'    Dim rp As Integer
'
'    Dim IsBorderRendered As Boolean
'
'    c = GetIDByName(mc.Name)
'    mg = MenuGrps(mc.Parent)
'    imgLHS = IIf(mc.LeftImage.NormalImage <> "", Preferences.ImgSpace, 0)
'    imgRHS = IIf(mc.RightImage.NormalImage <> "", Preferences.ImgSpace, 0)
'
'    If IsNS Then
'        leftP = mg.CmdsFXSize
'        topP = 0
'
'        DivWidth = DivWidth - 2 * leftP - 2 * mg.CmdsMarginX + CalcFrameBorder(mc.Parent, ascHorizontal)
'        CommandHeight = CommandHeight - 2 * mg.CmdsMarginY
'
'        table = "<layer" + _
'                    " left=" & leftP & _
'                    " top=" & topP & _
'                    " width=" & DivWidth & _
'                    " height=" & CommandHeight & _
'                ">"
'        With mc
'            If UseHoverImage Then
'                If IsFlash(.BackImage.HoverImage) Then
'                    table = table + CreateFlashCode(MenuCmds(c).BackImage.HoverImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.hBackColor), True)
'                End If
'            Else
'                If IsFlash(.BackImage.NormalImage) Then
'                    table = table + CreateFlashCode(MenuCmds(c).BackImage.NormalImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.nBackColor), True)
'                End If
'            End If
'
'            If .LeftImage.NormalImage <> "" Then
'                If UseHoverImage Then
'                    imgFile = ImgAbsPath + GetFileName(.LeftImage.HoverImage)
'                Else
'                    imgFile = ImgAbsPath + GetFileName(.LeftImage.NormalImage)
'                End If
'                lImage = "<layer top=" & CInt((CommandHeight - .LeftImage.h) / 2) & " left=0>"
'                lImage = lImage + "<img src=""" + imgFile + """ width=" & .LeftImage.w & " height=" & .LeftImage.h & ">"
'                lImage = lImage + "</layer>"
'            End If
'
'            If .Caption <> "" Then
'                txt = "<layer top=" & CInt((CommandHeight - GetCmdTextHeight(c)) / 2)
'                lp = .LeftImage.w + imgLHS
'                rp = .RightImage.w + imgRHS
'                txt = txt + " width=" & DivWidth - (lp + rp) & " left=" & lp & ">"
'                Select Case .Alignment
'                    Case tacLeft
'                        txt = txt + "<div align=left>"
'                    Case tacCenter
'                        txt = txt + "<div align=center>"
'                    Case tacRight
'                        txt = txt + "<div align=right>"
'                End Select
'                txt = txt + PrepareCaption(.Caption, True) + "</div></layer>"
'            End If
'
'            If .RightImage.NormalImage <> "" Then
'                If UseHoverImage Then
'                    imgFile = ImgAbsPath + GetFileName(.RightImage.HoverImage)
'                Else
'                    imgFile = ImgAbsPath + GetFileName(.RightImage.NormalImage)
'                End If
'                rImage = "<layer top=" & CInt((CommandHeight - .RightImage.h) / 2) & " left=" & DivWidth - .RightImage.w & ">"
'                rImage = rImage + "<img src=""" + imgFile + """ width=" & .RightImage.w & " height=" & .RightImage.h & ">"
'                rImage = rImage + "</layer>"
'            End If
'
'            If lImage <> "" Or txt <> "" Or rImage <> "" Then
'                table = table + lImage + txt + rImage + "</layer>"
'            Else
'                table = ""
'            End If
'        End With
'    Else
'        If UseHoverImage Then
'            IsBorderRendered = (mg.CmdsFXhColor <> -2) Or (mg.CmdsFXOver <> cfxcNone)
'            Select Case mg.CmdsFXOver
'                Case cfxcNone
'                    leftP = IIf(IsBorderRendered, 0, mg.CmdsFXSize) + mg.CmdsMarginX
'                    topP = IIf(IsBorderRendered, 0, mg.CmdsFXSize) + mg.CmdsMarginY
'                Case Else
'                    leftP = mg.CmdsMarginX
'                    topP = mg.CmdsMarginY
'            End Select
'        Else
'            IsBorderRendered = (mg.CmdsFXnColor <> -2) Or (mg.CmdsFXNormal <> cfxcNone)
'            Select Case mg.CmdsFXNormal
'                Case cfxcNone
'                    leftP = IIf(IsBorderRendered, 0, mg.CmdsFXSize) + mg.CmdsMarginX
'                    topP = IIf(IsBorderRendered, 0, mg.CmdsFXSize) + mg.CmdsMarginY
'                Case Else
'                    leftP = mg.CmdsMarginX
'                    topP = mg.CmdsMarginY
'            End Select
'        End If
'        If Not IsBorderRendered Then
'            If Int(mg.CmdsFXSize / 2) <> mg.CmdsFXSize / 2 Then
'                topP = topP - 1
'            End If
'        End If
'
'        DivWidth = DivWidth - 2 * leftP - IIf(IsBorderRendered, 2 * mg.CmdsFXSize, 0)
'        CommandHeight = CommandHeight - 2 * topP
'
'        table = "<span style=""position:absolute" + _
'                    "; top: " & topP + IIf(IsBorderRendered, 0, mg.CmdsFXSize) & "px" & _
'                    "; left: " & leftP & "px" & _
'                    "; width: " & DivWidth & "px" & _
'                    "; height: " & CommandHeight & "px" & _
'                ";"">"
'        With mc
'            If UseHoverImage Then
'                If IsFlash(.BackImage.HoverImage) Then
'                    table = table + CreateFlashCode(MenuCmds(c).BackImage.HoverImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.hBackColor), False)
'                End If
'            Else
'                If IsFlash(.BackImage.NormalImage) Then
'                    table = table + CreateFlashCode(MenuCmds(c).BackImage.NormalImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.nBackColor), False)
'                End If
'            End If
'
'            If .LeftImage.NormalImage <> "" Then
'                imgFile = ImgAbsPath + GetFileName(IIf(UseHoverImage, .LeftImage.HoverImage, .LeftImage.NormalImage))
'                lImage = "<span style=""position:absolute; top:" & CInt((CommandHeight - .LeftImage.h) / 2) & "px; left:0px;"">"
'                lImage = lImage + "<img src=""" + imgFile + """ width=" & .LeftImage.w & " height=" & .LeftImage.h & ">"
'                lImage = lImage + "</span>"
'            End If
'
'            If .Caption <> "" Then
'                txt = "<div style=""position:absolute; top:" & CInt((CommandHeight - GetCmdTextHeight(c)) / 2) & "px"
'                lp = .LeftImage.w + imgLHS
'                rp = .RightImage.w + imgRHS
'                txt = txt + "; width:" & DivWidth - (lp + rp) & "px; left:" & lp & "px;"""
'                Select Case .Alignment
'                    Case tacLeft
'                        txt = txt + " align=left"
'                    Case tacCenter
'                        txt = txt + " align=center"
'                    Case tacRight
'                        txt = txt + " align=right"
'                End Select
'                If UseHoverImage Then
'                    txt = txt + "></div>"
'                Else
'                    txt = txt + ">" + PrepareCaption(.Caption, False) + "</div>"
'                End If
'            End If
'
'            If .RightImage.NormalImage <> "" Then
'                imgFile = ImgAbsPath + GetFileName(IIf(UseHoverImage, .RightImage.HoverImage, .RightImage.NormalImage))
'                rImage = "<span style=""position:absolute; top:" & CInt((CommandHeight - .RightImage.h) / 2) & "px; left:" & DivWidth - .RightImage.w & "px;"">"
'                rImage = rImage + "<img src=""" + imgFile + """ width=" & .RightImage.w & " height=" & .RightImage.h & ">"
'                rImage = rImage + "</span>"
'            End If
'
'            table = table + lImage + txt + rImage + "</span>"
'        End With
'    End If
'
'    CreateImageTable = table
'
'End Function

Public Function PrepareCaption(ByVal c As String, IsNS As Boolean) As String

    c = xUNI2HTML(c)
    If IsNS Then c = NSCaptionFix(c)
    
'    c = Replace(c, ", ", ",&nbsp;")
'    c = Replace(c, ": ", ":&nbsp;")
'    c = Replace(c, "; ", ";&nbsp;")
'    c = Replace(c, "- ", "-&nbsp;")
'    c = Replace(c, " -", "&nbsp;-")
'    c = Replace(c, "/ ", "/&nbsp;")
'    c = Replace(c, " /", "&nbsp;/")

    Dim i As Integer
    Dim m As String
    
    i = 1
ReStart:
    For i = i To Len(c)
        m = Mid(c, i, 1)
        If Asc(m) >= 160 Then
            c = Left(c, i - 1) + "&#" & Asc(m) & Mid(c, i + 1)
            i = i + 3
            GoTo ReStart
        End If
    Next i

    PrepareCaption = c

End Function

Private Function NSCaptionFix(ByVal c As String) As String

    c = Replace(c, "|", "&#124")
    'c = Replace(c, "<", "&lt;")
    'c = Replace(c, ">", "&gt;")
    
    NSCaptionFix = c

End Function

Private Function CreateTBImageTable(g As Integer, UseHoverImage As Boolean, IsNS As Boolean, ByVal DivWidth As Integer, ByVal CommandHeight As Integer) As String

    Dim table As String
    Dim imgFile As String
    Dim imgLHS As Integer
    Dim imgRHS As Integer
    Dim mg As MenuGrp
    Dim leftP As Integer
    
    Dim lImage As String
    Dim rImage As String
    Dim txt As String
    Dim topP As Integer
    
    Dim lp As Integer
    Dim rp As Integer
    
    Dim IsBorderRendered As Boolean
    
    mg = MenuGrps(g)
    imgLHS = IIf(LenB(mg.tbiLeftImage.NormalImage) <> 0, Preferences.ImgSpace, 0)
    imgRHS = IIf(LenB(mg.tbiRightImage.NormalImage) <> 0, Preferences.ImgSpace, 0)
    
    If IsNS Then
        leftP = mg.CmdsFXSize
        topP = 0
        
        DivWidth = DivWidth - 2 * leftP - 2 * mg.CmdsMarginX
        CommandHeight = CommandHeight - 2 * mg.CmdsMarginY
        
        table = "<layer" + _
                    " left=" & leftP & _
                    " top=" & topP & _
                    " width=" & DivWidth & _
                    " height=" & CommandHeight & _
                ">"
        With mg
            If UseHoverImage Then
                If IsFlash(.tbiBackImage.HoverImage) Then
                    table = table + CreateFlashCode(.tbiBackImage.HoverImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.hBackColor, True), True)
                End If
            Else
                If IsFlash(.tbiBackImage.NormalImage) Then
                    table = table + CreateFlashCode(.tbiBackImage.NormalImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.nBackColor, True), True)
                End If
            End If
            
            If LenB(.tbiLeftImage.NormalImage) <> 0 Then
                If UseHoverImage Then
                    imgFile = ImgAbsPath + GetFileName(.tbiLeftImage.HoverImage)
                Else
                    imgFile = ImgAbsPath + GetFileName(.tbiLeftImage.NormalImage)
                End If
                lImage = "<layer top=" & CInt((CommandHeight - .tbiLeftImage.h) / 2) & " left=0>"
                lImage = lImage + "<img src=""" + imgFile + """ width=" & .tbiLeftImage.w & " height=" & .tbiLeftImage.h & ">"
                lImage = lImage + "</layer>"
            End If
            
            If LenB(.Caption) <> 0 Then
                txt = "<layer top=" & CInt((CommandHeight - GetGrpTextHeight(g)) / 2)
                lp = .tbiLeftImage.w + imgLHS
                rp = .tbiRightImage.w + imgRHS
                txt = txt + " width=" & DivWidth - (lp + rp) & " left=" & lp & ">"
                Select Case .CaptionAlignment
                    Case tacLeft
                        txt = txt + "<div align=""left"">"
                    Case tacCenter
                        txt = txt + "<div align=""center"">"
                    Case tacRight
                        txt = txt + "<div align=""right"">"
                End Select
                txt = txt + PrepareCaption(.Caption, True) + IIf(LenB(.Caption) = 0, "", "</div>") + "</layer>"
            End If
            
            If LenB(.tbiRightImage.NormalImage) <> 0 Then
                If UseHoverImage Then
                    imgFile = ImgAbsPath + GetFileName(.tbiRightImage.HoverImage)
                    'imgName = "name=" + .Name + "RImg"
                Else
                    imgFile = ImgAbsPath + GetFileName(.tbiRightImage.NormalImage)
                    'imgName = "name=" + .Name + "RImg"
                End If
                rImage = "<layer top=" & CInt((CommandHeight - .tbiRightImage.h) / 2) & " left=" & DivWidth - .tbiRightImage.w & ">"
                rImage = rImage + "<img src=""" + imgFile + """ width=" & .tbiRightImage.w & " height=" & .tbiRightImage.h & ">"
                rImage = rImage + "</layer>"
            End If
            
            table = table + lImage + txt + rImage + "</layer>"
        End With
    Else
        If UseHoverImage Then
            IsBorderRendered = (mg.CmdsFXhColor <> -2) Or (mg.CmdsFXOver <> cfxcNone)
            Select Case mg.CmdsFXOver
                Case cfxcNone
                    leftP = IIf(IsBorderRendered, 0, mg.CmdsFXSize) + mg.CmdsMarginX
                    topP = IIf(IsBorderRendered, 0, mg.CmdsFXSize) + mg.CmdsMarginY
                Case Else
                    leftP = mg.CmdsMarginX
                    topP = mg.CmdsMarginY
            End Select
        Else
            IsBorderRendered = (mg.CmdsFXnColor <> -2) Or (mg.CmdsFXNormal <> cfxcNone)
            Select Case mg.CmdsFXNormal
                Case cfxcNone
                    leftP = IIf(IsBorderRendered, 0, mg.CmdsFXSize) + mg.CmdsMarginX
                    topP = IIf(IsBorderRendered, 0, mg.CmdsFXSize) + mg.CmdsMarginY
                Case Else
                    leftP = mg.CmdsMarginX
                    topP = mg.CmdsMarginY
            End Select
        End If
        If Not IsBorderRendered Then
            If Int(mg.CmdsFXSize / 2) <> mg.CmdsFXSize / 2 Then
                topP = topP - 1
            End If
        End If
        
        DivWidth = DivWidth - 2 * leftP - IIf(IsBorderRendered, 2 * mg.CmdsFXSize, 0)
        CommandHeight = CommandHeight - 2 * mg.CmdsFXSize
        
        With mg
            If UseHoverImage Then
                If IsFlash(.tbiBackImage.HoverImage) Then
                    table = table + CreateFlashCode(.tbiBackImage.HoverImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.hBackColor, True), False)
                End If
            Else
                If IsFlash(.tbiBackImage.NormalImage) Then
                    table = table + CreateFlashCode(.tbiBackImage.NormalImage, CStr(DivWidth), CStr(CommandHeight), GetRGB(.nBackColor, True), False)
                End If
            End If
            
            If LenB(.tbiLeftImage.NormalImage) <> 0 Then
                imgFile = ImgAbsPath + GetFileName(IIf(UseHoverImage, .tbiLeftImage.HoverImage, .tbiLeftImage.NormalImage))
                lImage = lImage + "<img" + _
                                    IIf(UsesDynAPI, " id=""L"" ", " ") + _
                                    "style=""position:absolute; top:" & CInt((CommandHeight - .tbiLeftImage.h) / 2) + IIf(IsBorderRendered, 0, mg.CmdsFXSize) & "px; left:" & leftP + .tbiLeftImage.Margin & "px;"" src=""" + imgFile + """ width=" & .tbiLeftImage.w & " height=" & .tbiLeftImage.h & ">"
            End If
            
            If LenB(.Caption) <> 0 Then
                txt = "<div style=""position:absolute; top:" & CInt((CommandHeight - GetGrpTextHeight(g)) / 2) + IIf(IsBorderRendered, 0, mg.CmdsFXSize) & "px"
                lp = .tbiLeftImage.w + imgLHS
                rp = .tbiRightImage.w + imgRHS
                txt = txt + "; width:" & DivWidth - (lp + rp) & "px; left:" & lp + leftP & "px;"" "
                Select Case .CaptionAlignment
                    Case tacLeft
                        txt = txt + "align=""left"""
                    Case tacCenter
                        txt = txt + "align=""center"""
                    Case tacRight
                        txt = txt + "align=""right"""
                End Select
                If UseHoverImage Then
                    txt = txt + "></div>"
                Else
                    txt = txt + ">" + PrepareCaption(.Caption, False) + "</div>"
                End If
            End If
            
            If LenB(.tbiRightImage.NormalImage) <> 0 Then
                imgFile = ImgAbsPath + GetFileName(IIf(UseHoverImage, .tbiRightImage.HoverImage, .tbiRightImage.NormalImage))
                rImage = rImage + "<img" + _
                                    IIf(UsesDynAPI, " id=""R"" ", " ") + _
                                    "style=""position:absolute; top:" & CInt((CommandHeight - .tbiRightImage.h) / 2) + IIf(IsBorderRendered, 0, mg.CmdsFXSize) & "px; left:" & DivWidth - .tbiRightImage.w + leftP - .tbiRightImage.Margin & "px;"" src=""" + imgFile + """ width=" & .tbiRightImage.w & " height=" & .tbiRightImage.h & ">"
            End If
            
            table = table + lImage + txt + rImage '+ "</span>"
        End With
    End If
    
    CreateTBImageTable = table

End Function

Public Function EncodeStr(ByVal str As String) As String

    str = Replace(str, """", "\""")
    str = Replace(str, "'", "\'")
    str = Replace(str, vbCrLf, "")
    str = Replace(str, vbTab, "")
    
    str = RemoveSSEncoding(str, "%")
    str = RemoveSSEncoding(str, "?")
    
    EncodeStr = str

End Function

Private Function RemoveSSEncoding(ByVal sStr As String, lTag As String) As String

    Dim p1 As Long
    Dim p2 As Long
    Dim p As Long
    Dim oTag As String
    Dim cTag As String
    Dim nStr As String
    
    
    oTag = "<" + lTag
    cTag = lTag + ">"
    
    p1 = 1
    p2 = 1
    Do
        p1 = InStr(p1, sStr, oTag)
        If p1 = 0 Then Exit Do
        p2 = InStr(p1, sStr, cTag)
        p = InStr(p1, sStr, "\""")
        If p > p1 And p < p2 Then
            nStr = Replace(Mid(sStr, p1, p2 - p1), "\""", """")
            nStr = nStr + Space(Len(Mid(sStr, p1, p2 - p1)) - Len(nStr))
            Mid(sStr, p1, p2 - p1) = nStr
            p1 = p2
        Else
            Exit Do
        End If
    Loop
    
    RemoveSSEncoding = sStr

End Function

'Private Function ImageObj(mc As MenuCmd) As String
'
'    Dim f As String
'    Dim nimgFile As String
'    Dim himgFile As String
'    Dim fImgPath As String
'
'    If ImgAbsPath = "%%REL%%" Then
'        fImgPath = "'+rimPath+'"
'    Else
'        fImgPath = ImgAbsPath
'    End If
'
'    With mc
'        If .LeftImage.NormalImage <> "" Then
'
'            nimgFile = fImgPath + GetFileName(.LeftImage.NormalImage)
'            himgFile = fImgPath + GetFileName(.LeftImage.HoverImage)
'
'            f = String$(1, vbTab) + "var " + .Name + "LImgOff = new Image;" + vbCrLf
'            f = f + String$(1, vbTab) + "var " + .Name + "LImgOn = new Image;" + vbCrLf
'            f = f + String$(1, vbTab) + .Name + "LImgOff.src = '" + nimgFile + "';" + vbCrLf
'            f = f + String$(1, vbTab) + .Name + "LImgOn.src = '" + himgFile + "';" + vbCrLf
'        End If
'        If .RightImage.NormalImage <> "" Then
'
'            nimgFile = fImgPath + GetFileName(.RightImage.NormalImage)
'            himgFile = fImgPath + GetFileName(.RightImage.HoverImage)
'
'            f = f + String$(1, vbTab) + "var " + .Name + "RImgOff = new Image;" + vbCrLf
'            f = f + String$(1, vbTab) + "var " + .Name + "RImgOn = new Image;" + vbCrLf
'            f = f + String$(1, vbTab) + .Name + "RImgOff.src = '" + nimgFile + "';" + vbCrLf
'            f = f + String$(1, vbTab) + .Name + "RImgOn.src = '" + himgFile + "';" + vbCrLf
'        End If
'
'        If .BackImage.NormalImage <> "" Then
'
'            nimgFile = fImgPath + GetFileName(.BackImage.NormalImage)
'            himgFile = fImgPath + GetFileName(.BackImage.HoverImage)
'
'            f = f + String$(1, vbTab) + "var " + .Name + "BImgOff = new Image;" + vbCrLf
'            f = f + String$(1, vbTab) + "var " + .Name + "BImgOn = new Image;" + vbCrLf
'            f = f + String$(1, vbTab) + .Name + "BImgOff.src = '" + nimgFile + "';" + vbCrLf
'            f = f + String$(1, vbTab) + .Name + "BImgOn.src = '" + himgFile + "';" + vbCrLf
'        End If
'    End With
'
'    ImageObj = Replace(f, "''+", "")
'
'End Function

'Private Function OptimizeStyleCode(Code As String, OpStep As Integer) As String
'
'    Dim lines() As String
'    Dim NewLine As String
'    Dim i As Integer
'    Dim j As Integer
'    Dim Total As Integer
'
'    lines = Split(Code, vbCrLf)
'    Total = UBound(lines) - 1
'
'    FloodPanel.Caption = "Optimizing Style Code..."
'    For i = Total To 1 Step -1
'
'        Select Case OpStep
'            Case 1
'                FloodPanel.Value = (Total - i) / (Total * 2) * 100
'            Case 2
'                FloodPanel.Value = 50 + (Total - i) / (Total * 2) * 100
'        End Select
'
'        For j = 0 To UBound(lines) - 1
'            If Split(lines(i), "]=")(1) = Split(lines(j), "]=")(1) And j <> i Then
'                NewLine = Split(lines(i), "]=")(0) + "]=" + Mid$(Split(lines(j), "]=")(0), 2) + "];"
'                If Len(NewLine) < Len(lines(i)) Then
'                    lines(i) = NewLine
'                    Exit For
'                End If
'            End If
'        Next j
'    Next i
'
'    OptimizeStyleCode = Join(lines, vbCrLf)
'
'End Function

'Private Function OptimizeImagesCode(str As String) As String
'
'    Dim imgs() As String
'    Dim i As Integer
'    Dim j As Integer
'
'    imgs = Split(str, vbCrLf)
'
'    For i = 2 To UBound(imgs)
'        If InStr(imgs(i), ".src") Then
'            For j = 2 To i - 1
'                If imgs(j) <> "" And imgs(i) <> "" Then
'                    If InStr(imgs(i), ".src") And Split(imgs(i), "=")(1) = Split(imgs(j), "=")(1) Then
'                        imgs(i - 2) = Split(imgs(i - 2), "=")(0) + "= " + Trim$(Mid$(Split(imgs(j - 2), "=")(0), 6)) + ";"
'                        imgs(i) = ""
'                    End If
'                End If
'            Next j
'        End If
'    Next i
'
'    For i = 0 To UBound(imgs) - 1
'        If imgs(i) <> "" Then
'            OptimizeImagesCode = OptimizeImagesCode + imgs(i) + vbCrLf
'        End If
'    Next i
'
'End Function

Private Function GetNullRelStuff(IsForPreviewing As Boolean) As String

    If ThisConfig.Type <> ctcCDROM Or IsForPreviewing Then
        GetNullRelStuff = "function _purl(u){return u;}function _fip(img){return img.src;}" + vbCrLf
    End If

End Function

Private Function RemoveReduntantFunctions(ByVal Code As String) As String

    Dim f() As String
    Dim fName As String
    Dim i As Long
    Dim p As Long
    
    If Project.AddIn.Name = "" Then GoTo ExitFcn
    
    f = Split(LoadFile(AppPath + "AddIns\" + Project.AddIn.Name + ".ext"), "function ")
    
    For i = UBound(f) To 1 Step -1
        p = InStr(f(i), "(")
        If p > 0 Then
            fName = Left$(f(i), p - 1)
            If fName = "PrepareEvents" Then Exit For
            RemoveFunction Code, fName
        End If
    Next i
    
ExitFcn:
    RemoveReduntantFunctions = Code

End Function

Private Function BuildCode(TargetPlatform As TargetPlatformConstants, IsForPreviewing As Boolean, ieTable As String, nsTable As String, ieStyleCode As String, nsStyleCode As String, bStyleCode As String, RenderBrowserCode As Boolean) As String

    Dim sStr As String
    Dim scriptCode As String
    Dim cLine As Integer
    Dim tLines As Integer
    Dim tbCode() As String
    Dim mOffset As String
    Dim lines() As String
    
    Const vbTab2 = vbTab + vbTab
    Const vbTab3 = vbTab + vbTab + vbTab
    Const vbTab4 = vbTab + vbTab + vbTab + vbTab
    
    If LenB(Project.AddIn.Name) <> 0 Then
        sStr = LoadFile(AppPath + "AddIns\" + Project.AddIn.Name + ".ext")
        sStr = Replace(sStr, ")0/ {", ") {")
        sStr = Replace(sStr, ")1/ {", ") {")
        lines = Split(sStr, vbCrLf)
        tLines = UBound(lines)
        For cLine = 0 To tLines
            If lines(cLine) = "***" Then
                cLine = cLine + 1
                Exit For
            End If
        Next cLine
    Else
        lines = Split(LoadFile(AppPath + "rsc\code.dat"), vbCrLf)
        tLines = UBound(lines)
        cLine = 0
    End If
    
    tbCode = BuildToolbar(IsForPreviewing, TargetPlatform)
    tbCode(1) = RemoveReduntantFunctions(tbCode(1))
    
    FloodPanel.Value = 0
    FloodPanel.Caption = "Creating Menus for " + _
                        IIf(TargetPlatform = tpcBoth, "all supported browsers", _
                            IIf(TargetPlatform = tpcIE, "DOM compliant browsers", "Navigator 4"))
    If LenB(Project.AddIn.Name) <> 0 Then FloodPanel.Caption = FloodPanel.Caption + " using AddIn " + Project.AddIn.Name
    
    If IsMultiProject Then
        For cLine = cLine To tLines
            If InStr(lines(cLine), "%%TOOLBARVARS") Then Exit For
        Next cLine
    End If
    
    For cLine = cLine To tLines
        FloodPanel.Value = CInt(cLine / tLines * 100)
        sStr = lines(cLine)
        Do While InStr(sStr, vbTab + "%%")
            sStr = Replace(sStr, vbTab + "%%", "%%")
        Loop
        If Left(sStr, 2) = "%%" Then
            Select Case sStr
                Case "%%COPYRIGHT"
                    If Not IsForPreviewing And Project.CodeOptimization = cocDEBUG Then
                        scriptCode = scriptCode + "//////////////////////////////////////////////////" + vbCrLf
                        scriptCode = scriptCode + GenComment("DMB DHTML ENGINE", EngineVersion)
                        scriptCode = scriptCode + GenComment("(c)xFX JumpStart", "")
                        scriptCode = scriptCode + GenComment("", "")
                        scriptCode = scriptCode + GenComment("PSN:", GenLicense)
                        scriptCode = scriptCode + GenComment("", "")
                        scriptCode = scriptCode + GenComment("GENERATED:", Date & " - " & Time)
                        scriptCode = scriptCode + GenComment("--------------------------------------------", "")
                        scriptCode = scriptCode + GenComment(" Config:", Project.UserConfigs(Project.DefaultConfig).Name)
                        scriptCode = scriptCode + GenComment("  AddIn:", Project.AddIn.Name)
                        scriptCode = scriptCode + GenComment("JS Name:", Project.JSFileName)
                        scriptCode = scriptCode + "//////////////////////////////////////////////////" + vbCrLf
                    End If
                Case "%%FX"
                    If Not IsMultiProject Then
                        scriptCode = scriptCode + vbTab
                        If Project.FX < 4 Then
                            scriptCode = scriptCode + "var fx = " & Project.FX & ";" + vbCrLf + vbCrLf
                        Else
                            scriptCode = scriptCode + "var fx = Math.round(Math.random()*3);" + vbCrLf + vbCrLf
                        End If
                    End If
                Case "%%BROWSERCODE"
                    scriptCode = scriptCode + GetNullRelStuff(IsForPreviewing)
                    If RenderBrowserCode Then
                        If (TargetPlatform = tpcBoth Or TargetPlatform = tpcIE) And (Project.CompileIECode Or IsForPreviewing) Then
                            scriptCode = scriptCode + GenBrowserSaveCode(ieTable, "IE", IsForPreviewing)
                        End If
                        If ((TargetPlatform = tpcBoth Or TargetPlatform = tpcNS) And Not (IsForPreviewing And Val(GetSetting("DMB", "Browsers", "Default", 1)) = 1)) And (Project.CompileNSCode Or IsForPreviewing) Then
                            scriptCode = scriptCode + GenBrowserSaveCode(nsTable, "NS", IsForPreviewing)
                        End If
                    End If
                    scriptCode = scriptCode + "SetUpEvents();"
                Case "%%STYLECODE"
                    Select Case TargetPlatform
                        Case tpcBoth
                            scriptCode = scriptCode + ieStyleCode + nsStyleCode
                        Case tpcIE
                            If Project.CompileIECode Or IsForPreviewing Then scriptCode = scriptCode + ieStyleCode
                        Case tpcNS
                            If Project.CompileNSCode Then scriptCode = scriptCode + nsStyleCode
                    End Select
                    scriptCode = scriptCode + bStyleCode
                Case "%%BLINKFX"
                    scriptCode = scriptCode + vbTab + "var mibc = " & Project.BlinkEffect & ";" + vbCrLf
                    scriptCode = scriptCode + vbTab + "var mibm;" + vbCrLf
                    scriptCode = scriptCode + vbTab + "var mibs = " & Project.BlinkSpeed & ";"
                Case "%%BLINKDIS"
                    If Project.BlinkEffect > 0 Then scriptCode = scriptCode + "if(mibm) return false;"
                Case "%%DXFILTER"
                    scriptCode = scriptCode + vbTab + "var dxFilter=" + IIf(LenB(Project.DXFilter) = 0, "null", """" + Project.DXFilter + """") + ";" + vbCrLf
                Case "%%FRAMESIG"
                    If FramesMode Then
                        scriptCode = scriptCode + vbTab + vbTab + vbTab + "cFrame.sig = """ + CleanProjectName + """;"
                    End If
'                Case "%%FRAMEEVENT"
'                    If FramesMode Then
'                        scriptCode = scriptCode + vbTab + "mFrame.document.onmousemove = HideMenus;" + vbCrLf
'                    End If
                Case "%%BODYOBJCODE%%"
                    Dim RealLocal As ConfigDef
                    RealLocal = IIf(Project.UserConfigs(Project.DefaultConfig).Type = ctcRemote, Project.UserConfigs(GetConfigID(Project.UserConfigs(Project.DefaultConfig).LocalInfo4RemoteConfig)), Project.UserConfigs(Project.DefaultConfig))
                    If RealLocal.Frames.UseFrames And (TargetPlatform = tpcIE) And (Val(GetSetting("DMB", "Preferences", "ForceIE4", 0)) = 0) Then
                        scriptCode = scriptCode + vbTab2 + "try {"
                    End If
                    scriptCode = scriptCode + vbTab2 + "if(!f.bo && f.document)" + vbCrLf
                    scriptCode = scriptCode + vbTab3 + "f.bo = (f.document.compatMode==""BackCompat""||BV<6||IsMac)?f.document.body:f.document.documentElement;" + vbCrLf
                    scriptCode = scriptCode + vbTab2 + "return f.bo;"
                    If RealLocal.Frames.UseFrames And (TargetPlatform = tpcIE) And (Val(GetSetting("DMB", "Preferences", "ForceIE4", 0)) = 0) Then
                        scriptCode = scriptCode + vbTab2 + "} catch(e) {IsFrames=false;mFrame=cFrame;return GetBodyObj(cFrame);}"
                    End If
                Case "%%TOOLBARVARS"
                    scriptCode = scriptCode + tbCode(0)
                Case "%%TOOLBARSTYLE"
                    scriptCode = scriptCode + tbCode(1)
                Case "%%TOOLBARCODE"
                    scriptCode = scriptCode + tbCode(2)
                Case "%%CONTEXTMENU"
                    scriptCode = scriptCode + vbTab + "var cntxMenu = '" + GetContextGroupName + "';"
                Case "%%MENUSOFFSET"
                    mOffset = ""
                    If Project.MenusOffset.RootMenusX <> 0 Then mOffset = mOffset + "xy[0] += " & Project.MenusOffset.RootMenusX & ";"
                    If Project.MenusOffset.RootMenusY <> 0 Then mOffset = mOffset + "xy[1] += " & Project.MenusOffset.RootMenusY & ";"
                    If LenB(mOffset) <> 0 Then mOffset = vbTab3 + "if(!IsContext) {" + mOffset + "}" + vbCrLf
                    scriptCode = scriptCode + mOffset
                    scriptCode = scriptCode + GetCustomOffset("RootMenus")
                Case "%%SUBMENUSOFFSET"
                    If Project.MenusOffset.SubMenusX <> 0 Then scriptCode = scriptCode + vbTab4 + "xy[0] += " & Project.MenusOffset.SubMenusX & ";" + vbCrLf
                    If Project.MenusOffset.SubMenusY <> 0 Then scriptCode = scriptCode + vbTab4 + "xy[1] += " & Project.MenusOffset.SubMenusY & ";" + vbCrLf
                    scriptCode = scriptCode + GetCustomOffset("SubMenus")
                Case "%%FORMSTWEAK"
                    scriptCode = scriptCode + vbTab + "var DoFormsTweak = " + IIf(Project.DoFormsTweak, "true", "false") + ";"
                Case "%%TIMERS"
                    scriptCode = scriptCode + vbTab + "var AnimSpeed = " & Project.AnimSpeed & ";" & vbCrLf
                    scriptCode = scriptCode + vbTab + "var TimerHideDelay = " & Project.HideDelay & ";" & vbCrLf
                    scriptCode = scriptCode + vbTab + "var smDelay = " & Project.SubMenusDelay & ";" & vbCrLf
                    scriptCode = scriptCode + vbTab + "var rmDelay = " & Project.RootMenusDelay & ";" & vbCrLf
                    scriptCode = scriptCode + vbTab + "var scDelay = " & Project.SelChangeDelay & ";" & vbCrLf
                Case "%%DWSUPPORT01"
                    If Project.DWSupport Then
                        scriptCode = scriptCode + vbTab + "var DMB_MM_swapImgRestore;" & vbCrLf
                        scriptCode = scriptCode + vbTab + "var DMB_MM_nbGroup;" & vbCrLf
                    End If
                Case "%%DWSUPPORT02"
                    If Project.DWSupport Then
                        scriptCode = scriptCode + vbTab3 + "if(nOM==0){if(DMB_MM_swapImgRestore)DMB_MM_swapImgRestore();if(DMB_MM_nbGroup){om[1].state = 1;DMB_MM_nbGroup('out');}}" & vbCrLf
                    End If
                Case "%%DWSUPPORT03"
                    If Project.DWSupport Then
                        scriptCode = scriptCode + vbTab2 + "if(!DMB_MM_swapImgRestore&&typeof(MM_swapImgRestore)!='undefined'){DMB_MM_swapImgRestore = MM_swapImgRestore;MM_swapImgRestore = new Function('if(nOM==0)DMB_MM_swapImgRestore();');}" & vbCrLf
                        scriptCode = scriptCode + vbTab2 + "if(!DMB_MM_nbGroup&&typeof(MM_nbGroup)!='undefined'){DMB_MM_nbGroup = MM_nbGroup;MM_nbGroup = myMMfcn;}" & vbCrLf
                    End If
                Case "%%DWSUPPORT04"
                    If Project.DWSupport Then
                        scriptCode = scriptCode + vbTab + "function myMMfcn(){if(nOM>0){m=om[1];if(m.state==0 || typeof(m.state)=='undefined')m.state=1;if(m.state==1){m.state=2;var f = 'DMB_MM_nbGroup(';for(var i=0; i<arguments.length; i++){switch(typeof(arguments[i])){case 'string':f += ""'"" + arguments[i] + ""'"";break;case 'number':f += arguments[i].toString();break;}if(i<arguments.length-1) f += ',';}f += ')';eval(f);}}}" & vbCrLf
                    End If
                Case "%%NS4CLIPBUG"
                    If Not Project.NS4ClipBug And (TargetPlatform = tpcBoth Or TargetPlatform = tpcNS) Then
                        scriptCode = scriptCode + vbTab3 + "Menu.clip.width = 0;" & vbCrLf
                        scriptCode = scriptCode + vbTab3 + "Menu.clip.height = 0;" & vbCrLf
                    End If
'                Case "%%OPSUPPORT"
'                    If Project.OPHelperFunctions Then
'                        If TargetPlatform = tpcBoth Or TargetPlatform = tpcIE Then
'                            scriptCode = scriptCode + LoadFile(AppPath + "rsc\ophf.dat")
'                        End If
'                    Else
'                        scriptCode = scriptCode + vbTab + "function GetOPStyle(){;}function SetOPStyle(){;}"
'                    End If
                Case "%%IRSUPPORT01"
                    If Project.ImageReadySupport Then
                        scriptCode = scriptCode + vbTab + "function DMB_changeImages(){if(arguments.length==1)arg=arguments[0].toString().split(',');else  arg=arguments;for(var i=0;i<arg.length;i++)arg[i]=xrep(arg[i],""'"",'');if(document.images&&(preloadFlag==true)){for(i=0;i<arg.length;i+=2){if(document[arg[i]])document[arg[i]].src=arg[i+1];}}}" & vbCrLf
                    End If
                Case "%%IRSUPPORT02"
                    If Project.ImageReadySupport Then
                        scriptCode = scriptCode + vbTab3 + "if(nOM==1)if(om[nOM].irp)DMB_changeImages(om[nOM].irp);" & vbCrLf
                    End If
                Case "%%IRSUPPORT03"
                    If Project.ImageReadySupport Then
                        scriptCode = scriptCode + vbTab2 + "changeImages=new Function(""var arg=\""'\"";var ar=[];for(var i=0;i<arguments.length;i++)arg=arg+arguments[i]+\""','\"";arg=arg.substr(0,arg.length-2);ar=arg.split(','); if(nOM>0){if(om[nOM].irp!=true){DMB_changeImages(ar);om[nOM].irp=true;}else om[nOM].irp=ar;}else DMB_changeImages(ar);"");" & vbCrLf
                    End If
                Case "%%BROWSERDETECTION"
                    scriptCode = scriptCode + LoadFile(AppPath + "rsc\bdet.dat")
                    If (TargetPlatform = tpcIE) And (Val(GetSetting("DMB", "Preferences", "ForceIE4", 0)) = 0) Then
                        scriptCode = scriptCode + "if(IE&&!SM&&(BV>5))try{document.execCommand(""BackgroundImageCache"", false, true);}catch(x){;};" + vbCrLf
                    End If
                Case "%%SCROLLOFFSET"
                    scriptCode = scriptCode + vbTab3 + "var gs = om[nOM].gs; if(gs) xy[1] += parseInt(gs.top);"
                Case "%%AUTOSELP1", "%%AUTOSELP3"
                    If Project.AutoSelFunction Then
                        scriptCode = scriptCode + vbTab3 + "if(mc.disable==true) return false;"
                    End If
                Case "%%KBDNAVSUP"
                    If Project.KeyboardSupport And (TargetPlatform = tpcIE Or TargetPlatform = tpcBoth) Then
                        scriptCode = scriptCode + LoadFile(AppPath + "rsc/kbdctrl.dat")
                    End If
                Case "%%AUTOSCROLL"
                    If TargetPlatform = tpcIE Or TargetPlatform = tpcBoth Or IsForPreviewing Then
                        If Project.AutoScroll.MaxHeight <> 0 Then
                            scriptCode = scriptCode + vbTab2
                            If TargetPlatform = tpcBoth Then scriptCode = scriptCode + "if(!NS) {"
                            scriptCode = scriptCode + "mFrame.RemScroller(mName);if(y+parseInt(Menu.style.height) > pWH[1]) mFrame.AddScroller(mName, pWH[1] - y, true, true);"
                            If TargetPlatform = tpcBoth Then scriptCode = scriptCode + "}"
                        End If
                    End If
                Case "%%PRINTTBS"
                    Dim t As Integer
                    Dim s As String
                    If IsMasterProject Then
                        s = "window.setTimeout("""
                        For t = 1 To 99
                            s = s + "if(typeof(IEpTB" & t & ")!='undefined')IEpTB" & t & "();"
                        Next t
                        s = s + """, 500);" + vbCrLf
                    Else
                        For t = 1 + tbBaseIndex To UBound(Project.Toolbars) + tbBaseIndex
                            If Project.Toolbars(t - tbBaseIndex).Alignment = tbacFree Then
                                s = s + vbTab4
                                Select Case TargetPlatform
                                    Case tpcIE
                                        s = s + "IEpTB" & t & "();" + vbCrLf
                                    Case tpcNS
                                        's = s + "NSpTB" & t & "();" + vbCrLf
                                    Case Else
                                        s = s + "if(IE)IEpTB" & t & "();" + vbCrLf + vbTab3 + _
                                                "" '"if(NS)NSpTB" & t & "();" + vbCrLf
                                End Select
                            End If
                        Next t
                    End If
                    If s <> "" Then s = vbTab3 + "if(typeof(cFrame.prtb)=='undefined') {" + vbCrLf + _
                                        vbTab4 + "cFrame.prtb = true;" + vbCrLf + _
                                        s + vbCrLf + _
                                        vbTab3 + "SetupToolbar();" + vbCrLf + _
                                        vbTab3 + "}" + vbCrLf
                    scriptCode = scriptCode + s
            End Select
            DoEvents
        Else
            scriptCode = scriptCode + removeJSComment(sStr)
'            If InStr(Left(sStr, 4), "//") Then
'                If Project.CodeOptimization = cocDEBUG And Not IsForPreviewing Then
'                    scriptCode = scriptCode + sStr
'                End If
'            Else
'                If InStr(sStr, "//") Then sStr = Left$(sStr, InStr(sStr, "//") - 1)
'                scriptCode = scriptCode + sStr
'            End If
        End If
        
        If Right$(scriptCode, 1) <> vbCrLf Then
            scriptCode = scriptCode + vbCrLf
        End If
    Next cLine
    
    If Project.AutoSelFunction Then scriptCode = scriptCode + vbCrLf + LoadFile(AppPath + "rsc\autosel.dat")
    
    ' Disable the RootMenus Delay to avoid conflicts
    If Project.DWSupport Or Project.ImageReadySupport Or Project.RootMenusDelay <= 1 Then
        scriptCode = Replace(scriptCode, "smHnd = window.setTimeout(f[1]+f[0], f[2]);", "if(isc)smHnd = window.setTimeout(f[1]+f[0], f[2]); else eval(f[1]+f[0]);")
    End If
    
    If Project.LotusDominoSupport And Not IsForPreviewing Then
        ModifyImagesPathsForLotusDomino scriptCode
    End If
    
    If Project.BlinkEffect = 0 Then
        scriptCode = Replace(scriptCode, "window.setTimeout(""execURL2('"" + escape(_purl(url)) + ""', '"" + tframe + ""')"", d);", "execURL2(escape(_purl(url)), tframe);")
    End If
    
    scriptCode = Replace(scriptCode, ":Transparent", ":transparent")
    
ExitFcn:
        
    BuildCode = scriptCode

End Function

Private Function removeJSComment(ByVal Code As String) As String

    If InStr(Code, "//") And Project.CodeOptimization <> cocDEBUG Then
        Code = Left$(Code, InStr(Code, "//") - 1)
    End If
    
    removeJSComment = Code

End Function

Private Function GetCustomOffset(objName As String) As String

    Dim s() As String
    Dim o() As String
    Dim si As Integer
    Dim c As String
    Dim idx As Integer
    Dim nidx As Integer
    Dim r As String
    
    On Error GoTo IgnoreOffsets
    
    If LenB(Project.CustomOffsets) = 0 Then Exit Function
    
    s = Split(Project.CustomOffsets, "`")
    For si = 1 To UBound(s)
        c = ""
        o = Split(s(si), "@")
        Select Case o(1)
            Case "W"
                Select Case objName
                    Case "Toolbars"
                        nidx = 9
                    Case "RootMenus"
                        nidx = 11
                    Case "SubMenus"
                        nidx = 13
                End Select
                
                Select Case Split(o(2), ":")(1)
                    Case 1: c = c + "if(IsWin)"
                    Case 2: c = c + "if(!IsWin && !IsMac)"
                    Case 3: c = c + "if(IsMac)"
                End Select
                c = c + vbCrLf
                
                c = c + "if("
                c = c + getBrowserVal(o(3)) + "&&"
                c = c + getBrowserVal(o(4)) + "&&"
                c = c + getBrowserVal(o(5)) + "&&"
                c = c + getBrowserVal(o(6)) + "&&"
                c = c + getBrowserVal(o(7)) + "&&"
                c = c + ")" + vbCrLf
                c = Replace(c, "true&&", "")
                c = Replace(c, "&&)", ")")
                
                If Split(o(8), ":")(1) > 0 Then
                    c = c + "if(BV"
                    Select Case Split(o(8), ":")(1)
                        Case 1: c = c + "=="
                        Case 2: c = c + ">"
                        Case 3: c = c + ">="
                        Case 4: c = c + "<"
                        Case 5: c = c + "<="
                        Case 6: c = c + "!="
                    End Select
                    c = c + Split(o(8), ":")(2) + ")"
                End If
                
                c = c + "{" + vbCrLf
                
                If objName = "Toolbars" Then
                    'c = c + "tbleft += " + Split(o(nidx), ":")(1) + "; tbtop += " + Split(o(nidx + 1), ":")(1) + ";" + vbCrLf
                    c = c + "tbl += " + Split(o(nidx), ":")(1) + "; tbt += " + Split(o(nidx + 1), ":")(1) + ";"
                Else
                    c = c + "xy[0] += " + Split(o(nidx), ":")(1) + "; xy[1] += " + Split(o(nidx + 1), ":")(1) + ";"
                End If
                
                c = c + vbCrLf
                c = c + "}"
            Case "A"
                Select Case objName
                    Case "Toolbars"
                        idx = 2
                        nidx = 5
                    Case "RootMenus"
                        idx = 3
                        nidx = 7
                    Case "SubMenus"
                        idx = 4
                        nidx = 9
                End Select
                
                If objName = "Toolbars" Then
                    c = Replace(Replace(Split(o(idx), ":")(1), "\n", vbCrLf), Chr(255), ":")
                    c = Replace(c, "%f", "{tbl += " + Split(o(nidx), ":")(1) + "; tbt += " + Split(o(nidx + 1), ":")(1) + ";}")
                Else
                    If LenB(o(idx)) <> 0 Then
                        c = Replace(Replace(Split(o(idx), ":")(1), "\n", vbCrLf), Chr(255), ":")
                        c = Replace(c, "%f", "{xy[0] += " + Split(o(nidx), ":")(1) + "; xy[1] += " + Split(o(nidx + 1), ":")(1) + ";}")
                    End If
                End If
        End Select
        r = r + c
    Next si
    
    GetCustomOffset = Replace(r, "if()", "")
    
    Exit Function
    
IgnoreOffsets:
    GetCustomOffset = "{alert('The Custom Offset code produced an invalid javascript function:\n\n" + EncodeStr(c) + "')}"

End Function

Private Function getBrowserVal(ByVal v As String) As String

    Dim c As String

    Select Case Split(v, ":")(1)
        Case 0: c = Split(v, ":")(0)
        Case 1: c = "!" + Split(v, ":")(0)
        Case 2: c = "true"
    End Select
    
    getBrowserVal = c

End Function

Private Sub LDImage(scriptCode As String, ByVal ImageName As String, Optional Reset As Boolean = False)

    Static fn() As String
    Dim i As Integer
    
    If Reset Then
        ReDim fn(0)
        Exit Sub
    End If
    
    If LenB(ImageName) = 0 Then Exit Sub
    ImageName = GetFileName(ImageName)
    
    For i = 1 To UBound(fn)
        If fn(i) = ImageName Then Exit Sub
    Next i
    ReDim Preserve fn(UBound(fn) + 1)
    fn(UBound(fn)) = ImageName
    
    scriptCode = Replace(scriptCode, ImageName, ImageName + "/$file/" + ImageName)
    
    DoEvents

End Sub

Private Sub ModifyImagesPathsForLotusDomino(scriptCode As String)

    Dim i As Integer
    
    LDImage "", "", True
    
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            LDImage scriptCode, .LeftImage.NormalImage
            LDImage scriptCode, .LeftImage.HoverImage
            
            LDImage scriptCode, .BackImage.NormalImage
            LDImage scriptCode, .BackImage.HoverImage
            
            LDImage scriptCode, .RightImage.NormalImage
            LDImage scriptCode, .RightImage.HoverImage
        End With
    Next i
    
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            LDImage scriptCode, .BackImage.NormalImage
            LDImage scriptCode, .BackImage.HoverImage
            
            With .CornersImages
                LDImage scriptCode, .gcBottomCenter
                LDImage scriptCode, .gcBottomLeft
                LDImage scriptCode, .gcBottomRight
                LDImage scriptCode, .gcLeft
                LDImage scriptCode, .gcRight
                LDImage scriptCode, .gcTopCenter
                LDImage scriptCode, .gcTopLeft
                LDImage scriptCode, .gcTopRight
            End With
            
            LDImage scriptCode, .Image
            
            LDImage scriptCode, .tbiBackImage.NormalImage
            LDImage scriptCode, .tbiBackImage.HoverImage
            
            LDImage scriptCode, .tbiLeftImage.NormalImage
            LDImage scriptCode, .tbiLeftImage.HoverImage
            
            LDImage scriptCode, .tbiRightImage.NormalImage
            LDImage scriptCode, .tbiRightImage.HoverImage
        End With
    Next i
    
    For i = 1 To UBound(Project.Toolbars)
        With Project.Toolbars(i)
            LDImage scriptCode, .Image
        End With
    Next i

End Sub

Private Function AddGroupsScrollingCode(IsForPreview As Boolean, IsFramesFile As Boolean) As String

    Dim g As Integer
    Dim actCode As String
    Dim sStyle As String
    Dim sCode As String
    Dim cFX() As String
    
    If Not IsForPreview Then
        If FramesMode And Not IsFramesFile Then Exit Function
    End If
    
    For g = 1 To UBound(MenuGrps)
        If MenuGrps(g).Compile Then
            With MenuGrps(g).Scrolling
                If .MaxHeight > 0 Then
                    cFX = GetBorderFX(g, True, True)
                    actCode = actCode + vbTab + "AddScroller(""" + MenuGrps(g).Name + """, " & .MaxHeight & ", " & IIf(.OnMouseOver, "true", "false") & ");" + vbCrLf
                    sStyle = sStyle + vbTab + "var nCS_" & MenuGrps(g).Name & " = ['" + GetRGB(.nColor, True) + "', '" + ImgAbsPath + GetFileName(.UpImage.NormalImage) + "', '" + ImgAbsPath + GetFileName(.DnImage.NormalImage) + "', " & .DnImage.w & ", " & .DnImage.h & ", " & .Margin & ", '" & cFX(1) & "', " & .FXSize & "];" + vbCrLf
                    sStyle = sStyle + vbTab + "var hCS_" & MenuGrps(g).Name & " = ['" + GetRGB(.hColor, True) + "', '" + ImgAbsPath + GetFileName(.UpImage.HoverImage) + "', '" + ImgAbsPath + GetFileName(.DnImage.HoverImage) & "', '" & cFX(2) & "', " & .FXSize & "];" + vbCrLf
                End If
            End With
        End If
    Next g
    
    With Project.AutoScroll
        If .MaxHeight <> 0 Then
            cFX = GetBorderFX(0, True, True, True)
            sStyle = sStyle + vbTab + "var nCS_g" & " = ['" + GetRGB(.nColor, True) + "', '" + ImgAbsPath + GetFileName(.UpImage.NormalImage) + "', '" + ImgAbsPath + GetFileName(.DnImage.NormalImage) + "', " & .DnImage.w & ", " & .DnImage.h & ", " & .Margin & ", '" & cFX(1) & "', " & .FXSize & "];" + vbCrLf
            sStyle = sStyle + vbTab + "var hCS_g" & " = ['" + GetRGB(.hColor, True) + "', '" + ImgAbsPath + GetFileName(.UpImage.HoverImage) + "', '" + ImgAbsPath + GetFileName(.DnImage.HoverImage) & "', '" & cFX(2) & "', " & .FXSize & "];" + vbCrLf
        End If
    End With
    
    If LenB(actCode) <> 0 Or Project.AutoScroll.MaxHeight <> 0 Then
        If IsMultiProject Then
            sCode = sStyle + vbCrLf + actCode
        Else
            sCode = LoadFile(AppPath + "rsc\dmbscrl.dat")
            sCode = Replace(sCode, "%%ARSTYLES", sStyle)
            sCode = Replace(sCode, "%%ACTCODE", actCode)
        End If
        If IsForPreview Or ThisConfig.Type = ctcRemote Then
            sCode = Replace(sCode, "rimPath", "''")
        End If
        AddGroupsScrollingCode = vbCrLf + sCode
    End If

End Function

Private Function CleanProjectName() As String

    Dim pName As String
    
    pName = Project.Name
    
    pName = Replace(pName, """", "")
    pName = Replace(pName, "'", "")
    pName = Replace(pName, "\", "")
    pName = Replace(pName, "/", "")
    pName = Replace(pName, " ", "")
    
    CleanProjectName = "DMB" + pName

End Function

Private Function GenDynAPICode(cCode As String) As String

    Dim i As Integer

    #If DEMO = 0 Then

        Dim Code As String
        Dim tmp As String
        Dim p1 As Long
        Dim p2 As Long
        
        Dim gID As Integer
        Dim cid As Integer
        
        Dim s() As String
    
        If Not Project.GenDynAPI Then Exit Function
        If Not FileExists(AppPath + "rsc\dynapi.dat") Then Exit Function
        
        gID = GetIDByName("DynAPI_Template")
        cid = GetIDByName("DynAPICmd")
    
        Code = "/* DynAPI 1.1 */" + vbCrLf
        Code = "var imgsep = " & Preferences.ImgSpace & ";" & vbCrLf
        
        Code = Code + "var tLeading = " & MenuGrps(gID).Leading & ";" + vbCrLf
        Code = Code + "var tSCHeight = " & GetSingleCommandHeight(cid) & ";" + vbCrLf
        Code = Code + "var ncID = " & (UBound(MenuCmds) - 1) & ";" + vbCrLf
        
        Code = Code + LoadFile(AppPath + "rsc\dynapi.dat")
        
        If Project.CodeOptimization <> cocDEBUG Then
            s = Split(Code, vbCrLf)
            For i = 0 To UBound(s)
                p1 = InStr(s(i), "//")
                If p1 > 0 Then
                    s(i) = Left(s(i), p1 - 1)
                End If
            Next i
            Code = Join(s, vbCrLf)
        End If
        
        If Project.CodeOptimization <> cocDEBUG Then Code = Replace(Code, ";;", ";true;")
        GenDynAPICode = Code
    
    #Else
        For i = 1 To UBound(MenuGrps)
            If MenuGrps(i).IsTemplate Then
                GenDynAPICode = "alert('DynAPI is only available on the registered version of DHTML Menu Builder');" + vbCrLf
                Exit For
            End If
        Next i
    #End If

End Function

Private Function GetContextGroupName() As String

    Dim g As Integer
    
    For g = 1 To UBound(MenuGrps)
        With MenuGrps(g)
            If .IsContext And .Compile Then
                GetContextGroupName = .Name
                Exit Function
            End If
        End With
    Next g

End Function

Private Function scriptExpand(Platform As TargetPlatformConstants, Code As String, IsFramesFile As Boolean) As String

    Dim sStr As String
    Dim i As Integer
    Dim j As Integer
    Dim c As String
    Dim p As Long
    Dim t As Integer
    
    j = 2
    t = UBound(ExpandableItems)
    If Platform = tpcIE Then
        If IsMultiProject Then
            sStr = vbCrLf + "function Expand" & tbBaseIndex & "(c){var x;"
        Else
            sStr = vbCrLf + "function Expand(c){var x;"
        End If
        For i = 1 To t
            FloodPanel.Value = 50 + (i / (t * 2)) * 100
            With ExpandableItems(i)
                If (.Platform = Platform Or .Platform = tpcBoth) And LenB(.sStr) <> 0 Then
                    c = GenCChar(j)
                    p = InStr(Code, c)
                    If p > 0 Then
                        sStr = sStr + "x=""" + .sStr + """;c=xrep(c,""" + c + """,x);"
                        'If .CanExpandArray And Not IsFramesFile Then
                        '    sStr = sStr + "xA(""" + c + """,x);"
                        'End If
                    End If
                    j = NextCompIndex(j)
                    DoEvents
                End If
            End With
        Next i
        sStr = sStr + "return c;"
        sStr = sStr + "}"
'        If Not IsFramesFile Then
'            sStr = sStr + "function xA(o,n){"
'            sStr = sStr + "for(var i=0;i<hS.length;i++){"
'            sStr = sStr + "hS[i]=xrep(hS[i],o,n);"
'            sStr = sStr + "}"
'            If CreateToolbar Then
'                sStr = sStr + "for(var i=1001;i<hshS.length;i++){"
'                sStr = sStr + "hshS[i]=xrep(hshS[i],o,n);"
'                sStr = sStr + "}"
'            End If
'            sStr = sStr + "}"
'        End If
    End If
    
    If Platform = tpcNS Then
        If IsMultiProject Then
            sStr = vbCrLf + "function Expand" & tbBaseIndex & "(c){"
        Else
            sStr = vbCrLf + "function Expand(c){"
        End If
        For i = 1 To t
            FloodPanel.Value = 50 + (i / (t * 2)) * 100
            With ExpandableItems(i)
                If (.Platform = Platform Or .Platform = tpcBoth) And LenB(.sStr) <> 0 Then
                    c = GenCChar(j)
                    p = InStr(Code, c)
                    If p > 0 Then
                        sStr = sStr + "code=xrep(code,""" + c + """,""" + .sStr + """);"
                        'If .CanExpandArray And Not IsFramesFile Then
                        '    sStr = sStr + "xA(""" + c + """,""" + .sStr + """);"
                        'End If
                    End If
                    j = NextCompIndex(j)
                    DoEvents
                End If
            End With
        Next i
        sStr = sStr + "return code;"
        sStr = sStr + "}"
        If Not IsFramesFile Then
            sStr = sStr + "function xA(o, n){"
            sStr = sStr + "for(var i=1;i<nTCode.length;i++){"
            sStr = sStr + "if(nTCode[i])nTCode[i]=xrep(nTCode[i],o,n);"
            sStr = sStr + "}"
            sStr = sStr + "}"
        End If
    End If
    
    If Platform = tpcBoth Then
        sStr = "function Expand(code){"
        sStr = sStr + "return code;"
        sStr = sStr + "}"
    End If
    
    scriptExpand = sStr

End Function

Private Sub SaveCode(ByVal FileName As String, scriptCode As String, hRefStatements() As String, IsForPreviewing As Boolean, TargetPlatform As TargetPlatformConstants, DoNotSave As Boolean, DynAPICode As String)

    Dim ff As Integer
    Dim g As Integer
    Dim CompiledPath As String
    Dim ImagesPath As String
    Dim IsFramesFile As Boolean
    
    Dim errMsg As String
    
    On Error Resume Next
    
    ApplyAddInParams scriptCode
    
    #If DEMO = 0 Then
        If Project.GenDynAPI Then
            scriptCode = Replace(scriptCode, _
                            "ShowMenu(mName, x, y, isc, hsimg, algn) {" + vbCrLf, _
                            "ShowMenu(mName, x, y, isc, hsimg, algn) {" + vbCrLf + vbTab + vbTab + "mFrame._bg(mFrame._getobj(mName));" + vbCrLf)
        End If
    #End If

    If IsForPreviewing Then
        scriptCode = Replace(scriptCode, String$(2, vbTab) + vbCrLf, "")
        SaveFile PreviewPath + FileName + ".js", scriptCode + AddGroupsScrollingCode(IsForPreviewing, IsFramesFile) + DynAPICode
    Else
        IsFramesFile = (InStr(FileName, "_frames") > 0)
        If Project.CodeOptimization = cocDEBUG Then
            scriptCode = scriptCode + AddGroupsScrollingCode(IsForPreviewing, IsFramesFile) + DynAPICode
        Else
            scriptCode = CompressCode(scriptCode, TargetPlatform, IsFramesFile) + AddGroupsScrollingCode(IsForPreviewing, IsFramesFile) + DynAPICode
            
            If Not IsFramesFile Then scriptCode = Replace(scriptCode, "SetUpEvents();", "Expand('');SetUpEvents();")
        End If
        
        If ThisConfig.Type = ctcRemote Then
            CompiledPath = Project.UserConfigs(GetConfigID(ThisConfig.LocalInfo4RemoteConfig)).CompiledPath
            ImagesPath = Project.UserConfigs(GetConfigID(ThisConfig.LocalInfo4RemoteConfig)).ImagesPath
        Else
            CompiledPath = ThisConfig.CompiledPath
            ImagesPath = ThisConfig.ImagesPath
        End If
        
        If IsMultiProject Then
            RemoveAllFunctions scriptCode, IsFramesFile
        Else
            RemoveUnneededFunctions scriptCode, IsFramesFile, TargetPlatform
        End If
        If Not IsFramesFile Then CompressFunctionsNames scriptCode, TargetPlatform
        
        If DoNotSave Then GoTo ExitSub
        
        If SaveFile(CompiledPath + FileName + ".js", scriptCode) Then
            If Project.UseGZIP And TargetPlatform = tpcIE And Project.CodeOptimization <> cocDEBUG Then
                Compress CompiledPath + FileName + ".js", scriptCode, True
                SaveFile CompiledPath + FileName + ".js", scriptCode
            End If
            CopyProjectImages ImagesPath
            If ThisConfig.Type = ctcCDROM Then
                FileCopy AppPath + "blank.gif", ThisConfig.ImagesPath + "dmb_i.gif"
                FileCopy AppPath + "blank.gif", ThisConfig.CompiledPath + "dmb_m.gif"
            End If
        Else
            GoTo ErrorSaving
            Exit Sub
        End If
    End If
    
    ff = FreeFile
    If IsForPreviewing Then
        GoTo ExitSub
'        If IsReadOnly(PreviewPath + "hRef.txt") Then
'            errMsg = "The file '" + PreviewPath + "hRef.txt' is read-only"
'            GoTo ErrorSaving
'        End If
'        Open PreviewPath + "hRef.txt" For Output As #ff
    Else
        If Not Project.CompilehRefFile Then GoTo ExitSub
        If IsReadOnly(CompiledPath + "\hRef.txt") Then
            errMsg = "The file '" + CompiledPath + "\hRef.txt' is read-only"
            GoTo ErrorSaving
        End If
        Open CompiledPath + "\hRef.txt" For Output As #ff
    End If
        For g = 1 To UBound(MenuGrps)
            If MenuGrps(g).Compile Then Print #ff, hRefStatements(g)
        Next g
        
        Print #ff, "# LOADER CODE. PLACE IT RIGHT AFTER THE <BODY> TAG"
        Print #ff, "# OF THE " + ThisConfig.HotSpotEditor.HotSpotsFile + " FILE"
        If ThisConfig.Frames.UseFrames Then
            Print #ff, "# IN THE FRAME WHERE THE HOTSPOTS OR TOOLBAR ARE LOCATED"
        End If
        Print #ff, GenLoaderCode(False, IsForPreviewing)
        
        If FramesMode Then
            Print #ff, "# FOR FRAMES SUPPORT YOU MUST ADD THE FOLLOWING LINE"
            Print #ff, "# TO EVERY PAGE THAT WILL DISPLAY THE MENUS"
            Print #ff, "# PLACE IT RIGHT AFTER THE <BODY> TAG"
            Print #ff, GenLoaderCode(True, IsForPreviewing)
        End If
    Close #ff
    
ExitSub:
    
    Exit Sub
    
ErrorSaving:
    MsgBox "An error has occurred saving the compiled files." + vbCrLf + vbCrLf + _
            "Error " & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "Error Compiling"

End Sub

Private Sub RemoveAllFunctions(sCode As String, IsFramesFile As Boolean)

    Dim i As Integer
    Dim f() As String
    Dim p As Long
    Dim fName As String
    Dim ExpandFName As String
    
    If IsMultiProject Then
        ExpandFName = "Expand" & tbBaseIndex
    Else
        ExpandFName = "Expand"
    End If

    p = InStr(sCode, "function")
    f = Split(sCode, vbCrLf)
    For i = 1 To UBound(f)
        If InStr(sCode, f(i)) > p Then Exit For
        If Left(f(i), 5) = vbTab + "var " Then
            If InStr(f(i), "tbNum") = 0 And InStr(f(i), "new Image") = 0 Then f(i) = ""
        End If
    Next i
    sCode = Join(f, vbCrLf)
    
    f = Split(sCode, "function")
    For i = 1 To UBound(f)
        p = InStr(2, f(i), "(") - 2
        If p > 0 Then
            fName = Mid(f(i), 2, p)
            If fName <> ExpandFName And fName <> "xA" And InStr(fName, " ") = 0 And InStr(fName, "IEpTB") = 0 Then
                If p > 1 Then RemoveFunction sCode, fName
            End If
        End If
    Next i
    
    Do Until (InStr(sCode, vbCrLf + vbCrLf) = 0)
        sCode = Replace(sCode, vbCrLf + vbTab + vbCrLf + vbTab, vbCrLf + vbTab)
        sCode = Replace(sCode, vbCrLf + vbCrLf, vbCrLf)
    Loop
    
    If IsFramesFile Then
        Project.DOMFramesCode = sCode
    Else
        Project.DOMCode = sCode
    End If

End Sub

Private Sub RemoveUnneededFunctions(sCode As String, IsFramesFile As Boolean, TargetPlatform As TargetPlatformConstants)

    Dim i As Integer
    Dim reqScroll As Boolean
    Dim reqAutoPos As Boolean
    Dim reqContext As Boolean
    
    If UBound(Project.SecondaryProjects) > 0 Then
        For i = 1 To UBound(Project.Toolbars)
            With Project.Toolbars(i)
                If .FollowHScroll Or .FollowVScroll Then reqScroll = True
                If .Alignment = 10 Then reqAutoPos = True
            End With
        Next i
        If Not reqScroll Then
            RemoveFunction sCode, "ScrollTB"
            sCode = Replace(sCode, "if(tbFollowHScroll[t] || tbFollowVScroll[t]) ScrollTB(lt);", "")
            sCode = Replace(sCode, "if(tbFollowHScroll[t]||tbFollowVScroll[t])ScrollTB(lt);", "")
        End If
        
        If Project.RemoveImageAutoPosCode Then
            If Not reqAutoPos Then
                RemoveFunction sCode, "AutoPos"
                RemoveFunction sCode, "GetImgXY"
                RemoveFunction sCode, "GetImgWH"
                RemoveFunction sCode, "getOffset"
                RemoveFunction sCode, "FindImage"
                RemoveFunction sCode, "GetImgOffset"
            End If
        End If
        
        For i = 1 To UBound(MenuGrps)
            With MenuGrps(i)
                If .IsContext And .Compile Then
                    reqContext = True
                    Exit For
                End If
            End With
        Next i
        If Not reqContext Then RemoveFunction sCode, "ShowContextMenu"
        If Project.BlinkEffect = 0 Then RemoveFunction sCode, "BlinkItem"
    End If
    
    Select Case TargetPlatform
        Case tpcIE
            RemoveFunction sCode, "NSHoverSel"
            RemoveFunction sCode, "hNSCClick"
            RemoveFunction sCode, "FindImage"
            RemoveFunction sCode, "GetImgOffset"
            RemoveFunction sCode, "FixPointSize"
            RemoveFunction sCode, "rHnd"
            RemoveFunction sCode, "PrepareEvents"
            RemoveFunction sCode, "hsNSHoverSel"
            If Project.FX = 0 Then
                RemoveFunction sCode, "Animate"
                sCode = Replace(sCode, "Animate()", "void(0)")
                sCode = Replace(sCode, "if(!IX&&!SM&&IsWin) clip=""rect(0 0 0 0)"";", "")
            End If
            If IsFramesFile Then
                Project.DOMFramesCode = sCode
            Else
                Project.DOMCode = sCode
            End If
        Case tpcNS
            RemoveFunction sCode, "GetCurCmd"
            RemoveFunction sCode, "HoverSel"
            RemoveFunction sCode, "FixCommands"
            RemoveFunction sCode, "cIF"
            RemoveFunction sCode, "GetBodyObj"
            RemoveFunction sCode, "GetBorderWidth"
            RemoveFunction sCode, "hsHoverSel"
            RemoveFunction sCode, "FixCursor"
            RemoveFunction sCode, "SwapMC"
            If IsFramesFile Then
                Project.NSFramesCode = sCode
            Else
                Project.NSCode = sCode
            End If
    End Select

End Sub

Private Sub ApplyAddInParams(Code As String)

    Dim i As Integer
    
    If LenB(Project.AddIn.Name) <> 0 Then
        For i = 1 To UBound(params)
            With params(i)
                Code = Replace(Code, "??" + .Name + "??", IIf(LenB(.Value) = 0, .Default, .Value))
            End With
        Next i
    End If

End Sub

Private Function OptimizeJSCode(ByVal Code As String) As String

    Dim cLen As Long
    
    If Code = "" Then Exit Function
    
    cLen = 0
    Do Until cLen = Len(Code)
        cLen = Len(Code)
        Code = Replace(Code, vbTab, "")
        Code = Replace(Code, " = ", "=")
        Code = Replace(Code, " == ", "==")
        Code = Replace(Code, " || ", "||")
        Code = Replace(Code, " && ", "&&")
        Code = Replace(Code, " < ", "<")
        Code = Replace(Code, " > ", ">")
        Code = Replace(Code, " + ", "+")
        Code = Replace(Code, " - ", "-")
        Code = Replace(Code, " += ", "+=")
        Code = Replace(Code, " -= ", "-=")
        Code = Replace(Code, " >= ", ">=")
        Code = Replace(Code, " <= ", "<=")
        Code = Replace(Code, ";" + vbCrLf, ";")
        Code = Replace(Code, ";;", ";")
        Code = Replace(Code, "; ", ";")
        Code = Replace(Code, ": ", ":")
        Code = Replace(Code, ", ", ",")
        Code = Replace(Code, "}" + vbCrLf, "}")

        Code = Replace(Code, "nStyle", "nS")
        Code = Replace(Code, "hStyle", "hS")
        Code = Replace(Code, "nLayer", "nL")
        Code = Replace(Code, "hLayer", "hL")
        Code = Replace(Code, "BrowserVersion", "BV")
        Code = Replace(Code, vbCrLf, " ")

        Code = Replace(Code, "{ ", "{")
        Code = Replace(Code, " {", "{")
        Code = Replace(Code, "} ", "}")
        Code = Replace(Code, " }", "}")
        
        Code = Replace(Code, ";status=\'\';", ";")
    Loop
    
    OptimizeJSCode = Code

End Function

Private Function CompressCode(Code As String, TargetPlatform As TargetPlatformConstants, IsFrameFile As Boolean) As String

    Dim i As Integer
    Dim j As Integer
    Dim sCode As String
    Dim eCode As String
    Dim n As Integer
    Dim c As String
    Dim t As Integer
    Dim GetNextChar As Boolean
    Dim imgCodePos As Long
    
    FloodPanel.Caption = "Compressing JavaScript Code..."
    FloodPanel.Value = 1
    
    ' Protect Strings
    For i = 1 To UBound(Project.Toolbars)
        With Project.Toolbars(i)
            If .Condition <> "" Then Code = Replace(Code, .Condition, "_TB:C" & PadZeros(Hex(i), 4))
        End With
    Next i
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            If .Caption <> "" And .Compile Then Code = Replace(Code, ">" + .Caption + "<", "_MC:" & PadZeros(Hex(i), 4))
        End With
    Next i
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            If .Caption <> "" And .Compile Then Code = Replace(Code, ">" + .Caption + "<", "_MG:" & PadZeros(Hex(i), 4))
        End With
    Next i
    
    ' Compress Code
    Code = OptimizeJSCode(Code)
    
    ' Unprotect Strings
    For i = 1 To UBound(Project.Toolbars)
        With Project.Toolbars(i)
            If .Condition <> "" Then Code = Replace(Code, "_TB:C" & PadZeros(Hex(i), 4), .Condition)
        End With
    Next i
    For i = 1 To UBound(MenuCmds)
        With MenuCmds(i)
            If .Caption <> "" And .Compile Then Code = Replace(Code, "_MC:" & PadZeros(Hex(i), 4), ">" + .Caption + "<")
        End With
    Next i
    For i = 1 To UBound(MenuGrps)
        With MenuGrps(i)
            If .Caption <> "" And .Compile Then Code = Replace(Code, "_MG:" & PadZeros(Hex(i), 4), ">" + .Caption + "<")
        End With
    Next i

    If IsFrameFile Then
        eCode = Left(Code, InStr(Code, "));close();}") - 1)
    Else
        Code = Replace(Code, "s.split(""cursor:url("")[1]", "s.split(""cursor:url("")[1]")
        If UBound(Project.Toolbars) > 0 Then
            sCode = Left(Code, InStr(Code, "function SetupToolbar"))
        Else
            sCode = Left(Code, InStr(Code, "function GetCurCmd"))
        End If
        imgCodePos = InStr(sCode, "LImgOff=new Image;")
        If imgCodePos = 0 Then imgCodePos = Len(sCode) + 1
        eCode = Mid(Code, InStr(Code, "function PrepareEvents"))
    End If
    
    j = 2
    t = UBound(ExpandableItems) - IIf(Len(ImgAbsPath) > 4, 1, 0)
    For i = 1 To t
        With ExpandableItems(i)
            FloodPanel.Value = (i / (2 * t)) * 100
            GetNextChar = False
            If .Platform = TargetPlatform Or TargetPlatform = tpcBoth Then
                c = GenCChar(j)
                'n = CountOccurrences(sCode, .sStr) + CountOccurrences(eCode, .sStr)
                n = CountOccurrences(eCode, .sStr)
                If n > 1 Then
                    eCode = Replace(eCode, .sStr, c)
                    GetNextChar = True
                End If
                If GetNextChar Then
                    j = NextCompIndex(j)
                    DoEvents
                Else
                    .sStr = ""
                End If
            End If
        End With
    Next i
    
    If IsFrameFile Then
        Code = eCode + Mid(Code, InStr(Code, "));close();}"))
    Else
        If UBound(Project.Toolbars) > 0 Then
            Code = sCode + Mid(Code, InStr(Code, "function SetupToolbar") + 1)
        Else
            Code = sCode + Mid(Code, InStr(Code, "function GetCurCmd") + 1)
        End If
        Code = Left(Code, InStr(Code, "function PrepareEvents") - 1) + eCode
    End If
    
    If Len(ImgAbsPath) > 4 Then Code = CompressPath(Code, ImgAbsPath, GenCChar(j))
    
    If IsMultiProject Then
        CompressCode = Code + scriptExpand(TargetPlatform, Code, IsFrameFile)
    Else
        CompressCode = Mid(Code, InStr(Code, "v")) + scriptExpand(TargetPlatform, Code, IsFrameFile)
    End If
    
    ' Compress the parseInt function
    CompressCode = Replace(CompressCode, "parseInt(", "pri(")
    If InStr(CompressCode, "function pri(n)") = 0 Then CompressCode = CompressCode + "function pri(n){return parseInt(n,10)}"

End Function

Private Sub CompressFunctionsNames(ByRef sCode As String, TargetPlatform As TargetPlatformConstants)

    Dim f() As String
    Dim i As Long
    Dim j As Long
    Dim nidx As Integer
    Dim tmp As String
    Dim vc As String
    
    If Project.CodeOptimization <> cocAggressive Then Exit Sub
    
    ' Remove unnecessary CASE statements
    #If DEVVER = 0 Then
        For i = 0 To 10
            If Not NeedsTBAlignmentCase(i) Then RemoveCase sCode, "SetupToolbar", CStr(i)
        Next i
        For i = 0 To 11
            If Not NeedsTBHSPosCase(i) Then RemoveCase sCode, "GetHSPos", CStr(i)
            If Not NeedsSubMenuPosCase(i) Then RemoveCase sCode, "GetSubMenuPos", CStr(i)
        Next i
    #End If
    
    If FramesMode Then Exit Sub
    
    nidx = 0
    f = Split(sCode, "function ")
    
    For i = 1 To UBound(f)
        f(i) = Left(f(i), InStr(f(i), "(") - 1)
    Next i
    
ReStart:
    For i = 1 To UBound(f)
        For j = 1 To UBound(f)
            If i <> j Then
                If Len(f(i)) > Len(f(j)) Then
                    tmp = f(i)
                    f(i) = f(j)
                    f(j) = tmp
                End If
            End If
        Next j
    Next i
    
    vc = " ;{}()?=+-*/!:""[,.'|&" + vbTab
    For i = 1 To UBound(f)
        If Len(f(i)) > 5 Then
            Select Case f(i)
                Case "ShowMenu"
                    ' do nothing
                Case "tHideAll", "execURL", "dmbHighlightItem", "dmbGOBC", "dmbHighlightTBItem", "dmbHighlightGroupItem"
                    ' do nothing
                Case Else
                    If (f(i) = "HideAll" And (Project.DWSupport Or Project.ImageReadySupport)) Or _
                       (Project.AutoSelFunction And InStr(f(i), "HoverSel") > 0) Or _
                       Left(f(i), 7) = "dmbAPI_" Or _
                       (InStr(f(i), vbNull) > 0) Then
                        ' do nothing
                    Else
                        For j = 1 To Len(vc)
                            sCode = Replace(sCode, Mid(vc, j, 1) + f(i) + "(", Mid(vc, j, 1) + "f" & nidx & "(")
                        Next j
                        sCode = Replace(sCode, vbCrLf + f(i) + ".", vbCrLf + "f" & nidx & ".")
                        For j = 1 To Len(vc)
                            sCode = Replace(sCode, Mid(vc, j, 1) + f(i) + ".", Mid(vc, j, 1) + "f" & nidx & ".")
                        Next j
                        
                        sCode = Replace(sCode, "(" + f(i) + ")", "(f" & nidx & ")")
                        sCode = Replace(sCode, " " + f(i) + "=", " f" & nidx & "=")
                        sCode = Replace(sCode, ")" + f(i) + "=", ")f" & nidx & "=")
                        sCode = Replace(sCode, "=" + f(i) + ";", "=f" & nidx & ";")
    
                        sCode = Replace(sCode, "= " + f(i) + ";", "= f" & nidx & ";")
                        sCode = Replace(sCode, "." + f(i) + " =", ".f" & nidx & "=")
                        sCode = Replace(sCode, "." + f(i) + "= ", ".f" & nidx & "=")
                        
                        sCode = Replace(sCode, "<" + f(i), "<f" & nidx)
                        sCode = Replace(sCode, "," + f(i) + ",", ",f" & nidx & ",")
                        
                        nidx = nidx + 1
                    End If
            End Select
        End If
    Next i
    
    sCode = Replace(sCode, "function  f", "function f")
    
    #If DEVVER = 0 Then
        sCode = Replace(sCode, "tbSpanning", "tbK")
        sCode = Replace(sCode, "tbWidth", "tbW")
        sCode = Replace(sCode, "tbHeight", "tbH")
        sCode = Replace(sCode, "tbBorder", "tbB")
        sCode = Replace(sCode, "tbStyle", "tbY")
    #End If
    
    If TargetPlatform = tpcIE Then
        Dim z As Long
        j = 1
        Do
            j = InStr(j, sCode, "if(NS){")
            If j = 0 Then Exit Do
            
            j = j + 7
            z = 1
            For i = j To Len(sCode)
                If Mid(sCode, i, 1) = "{" Then z = z + 1
                If Mid(sCode, i, 1) = "}" Then
                    z = z - 1
                    If z = 0 Then Exit For
                End If
            Next i
            sCode = Left(sCode, j - 1) + ";" + Mid(sCode, i)
        Loop
    End If

End Sub

Private Function NeedsSubMenuPosCase(ByVal c As Integer) As Boolean
    
    Dim cid As Integer
    
    For cid = 1 To UBound(MenuCmds)
        With MenuCmds(cid)
            If .Compile Then
                If .Actions.OnMouseOver.Type = atcCascade Then
                    If .Actions.OnMouseOver.TargetMenuAlignment = c Then
                        NeedsSubMenuPosCase = True
                        Exit Function
                    End If
                End If
                If .Actions.OnClick.Type = atcCascade Then
                    If .Actions.OnClick.TargetMenuAlignment = c Then
                        NeedsSubMenuPosCase = True
                        Exit Function
                    End If
                End If
            End If
        End With
    Next cid

End Function

Private Function NeedsTBHSPosCase(ByVal c As Integer) As Boolean

    Dim t As Integer
    Dim g As Integer
    
    For t = 1 To UBound(Project.Toolbars)
        For g = 1 To UBound(Project.Toolbars(t).Groups)
            If MenuGrps(GetIDByName(Project.Toolbars(t).Groups(g))).Alignment = c Then
                NeedsTBHSPosCase = True
                Exit Function
            End If
        Next g
    Next t

End Function

Private Function NeedsTBAlignmentCase(ByVal c As Integer) As Boolean
    
    Dim t As Integer
    
    For t = 1 To UBound(Project.Toolbars)
        If Project.Toolbars(t).Alignment = c Then
            NeedsTBAlignmentCase = True
            Exit Function
        End If
    Next t

End Function

Private Sub RemoveCase(s As String, f As String, c As String)

    Dim p1 As Long
    Dim p2 As Long
    Dim i As Long
    
    p1 = InStr(s, "function " + f)
    If p1 = 0 Then Exit Sub
    
    p1 = InStr(p1, s, "case " & c)
    If p1 = 0 Then Exit Sub
    
    p2 = InStr(p1, s, "break;")
    
    s = Left(s, p1 - 1) + Mid(s, p2 + Len("break;"))

End Sub

Public Sub RemoveFunction(ByRef sc As String, fcn As String)

    Dim p1 As Long
    Dim p2 As Long
    Dim i As Long
    Dim p As Long
    Dim k As Long
    
    fcn = "function " + fcn
    
    p = InStr(sc, fcn)
    If p > 0 Then
        For i = p To Len(sc)
            p1 = InStr(i, sc, "{")
            p2 = InStr(i, sc, "}")
            If p2 < p1 Or p1 = 0 Then
                k = k - 1
                i = p2
            Else
                i = p1
                k = k + 1
            End If
            If k = 0 Then Exit For
        Next i
        sc = Replace(sc, Mid(sc, p, i - p + 1), "")
    End If

End Sub

Private Function GenCChar(j As Integer) As String

    GenCChar = IIf(j > 28, cChar, "") + Chr(j)

End Function

Private Function NextCompIndex(j As Integer) As Integer

    j = j + 1
    If j = 10 Then j = 11
    If j = 13 Then j = 14
    If j = 34 Then j = 35
    If j = 39 Then j = 40
    If j = 41 Then j = 42
    If j = 92 Then j = 93
    
    NextCompIndex = j

End Function

Private Function CompressPath(Code As String, p As String, c As String) As String

    Dim p1 As Long
    
    Dim lp As String
    Dim rp As String

    p1 = InStr(Code, p)
    While p1 > 0
        lp = Left(Code, p1 - 1)
        If Not (Right(lp, 9) = "Off.src='" Or Right(lp, 8) = "On.src='") Then
            rp = Mid(Code, p1 + Len(p))
            Code = lp + c + rp
        End If
        p1 = InStr(p1 + 1, Code, p)
    Wend
    
    CompressPath = Code

End Function

Private Function CountOccurrences(c As String, s As String) As Integer

    On Error Resume Next
    CountOccurrences = UBound(Split(c, s))

End Function

Private Function GenBrowserSaveCode(str As String, BrowserVar As String, IsForPreviewing As Boolean, Optional SkipWriteCode As Boolean = False) As String

    Dim f As String
    Dim UseExpand As Boolean
    Dim cCode As String
    
    UseExpand = Project.CodeOptimization <> cocDEBUG And Not IsForPreviewing
    
    cCode = IIf(BrowserVar = "NS", "FixPointSize(", "") + IIf(UseExpand, "Expand" & IIf(IsMultiProject And Not IsForPreviewing, tbBaseIndex, "") & "(""", """") + EncodeStr(str) + IIf(UseExpand, """)", """") + IIf(BrowserVar = "NS", ")", "")
    If ThisConfig.Type = ctcCDROM And Not IsForPreviewing Then
        cCode = "xrep(" + cCode + ", '%' + '%REL%%', rimPath)"
    End If
    
    If SkipWriteCode Then
        f = cCode
    Else
        f = String$(1, vbTab) + "if(" + BrowserVar + ")" + vbCrLf
        f = f + String$(2, vbTab) + "with(document) {" + vbCrLf
        f = f + String$(3, vbTab) + "open();" + vbCrLf
        f = f + String$(3, vbTab) + "write(" + cCode + ");" + vbCrLf
        f = f + String$(3, vbTab) + "close();" + vbCrLf
        f = f + String$(2, vbTab) + "}" + vbCrLf
    End If
    
    GenBrowserSaveCode = f
    
End Function

Private Function GenBrowserTBSaveCode(str() As String, BrowserVar As String, IsForPreviewing As Boolean, t As Integer) As String

    Dim cs As String
    Dim cf As String
    
    Dim fn As String
    Dim jn As String
    
    Dim f As Integer
    
    f = UBound(Project.Toolbars) + tbBaseIndex
    For t = 1 + tbBaseIndex To f
        If LenB(str(t)) <> 0 Then
            If Project.Toolbars(t - tbBaseIndex).Alignment = tbacFree Then
                fn = BrowserVar + "pTB" + CStr(t) + "()"
                jn = "dmbTB" + CStr(t) + "ph"
                cf = cf + "function " + fn + " {"
                cf = cf + "if(!GetObj('" + jn + "')) window.setTimeout('" + fn + "', 10); else "
                If BrowserVar = "IE" Then
                    cf = cf + "GetObj('" + jn + "').innerHTML = "
                Else
                    cf = cf + "GetObj('" + jn + "').document.open();"
                    cf = cf + "GetObj('" + jn + "').document.write("
                End If
                cf = cf + GenBrowserSaveCode(str(t), BrowserVar, IsForPreviewing, True)
                If BrowserVar = "IE" Then
                    cf = cf + ";}" + vbCrLf
                Else
                    cf = cf + ");"
                    cf = cf + "GetObj('" + jn + "').document.close();"
                    If t = f Then cf = cf + "PrepareEvents();"
                    cf = cf + "}" + vbCrLf
                End If
            Else
                cs = cs + GenBrowserSaveCode(str(t), BrowserVar, IsForPreviewing) + vbCrLf
            End If
        End If
    Next t
    
    GenBrowserTBSaveCode = cs + cf

End Function

Private Function scriptExtraFramesCode(BrowserCode As String, IsForPreviewing As Boolean, TargetPlatform As TargetPlatformConstants) As String

    Dim fCode As String
    
    fCode = LoadFile(AppPath + "rsc\fcode.dat")
    fCode = Replace(fCode, "%%BROWSERCODE", BrowserCode)
    fCode = Replace(fCode, "%%BROWSERDETECTION", LoadFile(AppPath + "rsc\bdet.dat"))
    fCode = Replace(fCode, "%%PNAME", CleanProjectName)
    
    If (Not IsForPreviewing) And (TargetPlatform = tpcIE) And (Val(GetSetting("DMB", "Preferences", "ForceIE4", 0)) = 0) Then
        fCode = Replace(fCode, "%%BSEC01", "try {")
        fCode = Replace(fCode, "%%BSEC02", "} catch(er) {return false;}")
    Else
        fCode = Replace(fCode, "%%BSEC01", "")
        fCode = Replace(fCode, "%%BSEC02", "")
    End If
    
    scriptExtraFramesCode = GetNullRelStuff(IsForPreviewing) + fCode

End Function

Private Function BuildToolbar(IsForPreviewing As Boolean, TargetPlatform As TargetPlatformConstants) As String()

    Dim g As Integer
    Dim ieTable() As String
    Dim nsTable() As String
    Dim mc As MenuCmd
    Dim scriptCode As String
    Dim ieStyleCode As String
    Dim nsStyleCode As String
    Dim cStyle As String
    Dim i As Integer
    Dim hsx As Integer
    Dim hsy As Integer
    Dim CmdLayerName As String
    Dim nsEC As String
    Dim scHeight As Integer
    Dim totalGroupsInTB As Integer
    Dim cFX() As String
    Dim n As String
    Dim o As String
    Dim t As Integer
    Dim k As Integer
    Dim tbCode(0 To 2) As String
    Dim imgCode As String
    Dim totalWidth As Integer
    Dim totalHeight As Integer
    Dim Wider As Integer
    Dim Taller As Integer
    Dim r() As Integer
    Dim dw As String
    Dim dH As String
    Dim hsWidth As Integer
    Dim hsHeight As Integer
    Dim dispTips As Boolean
    
    If Not CreateToolbar Then GoTo SkipBuildTB
    
    dispTips = Project.StatusTextDisplay = socBoth Or Project.StatusTextDisplay = socTooltip
    
    ReDim ieTable(UBound(Project.Toolbars) + tbBaseIndex)
    ReDim nsTable(UBound(Project.Toolbars) + tbBaseIndex)

    k = gBaseIndex
    For t = 1 + tbBaseIndex To UBound(Project.Toolbars) + tbBaseIndex
        
        Project.Toolbar = Project.Toolbars(t - tbBaseIndex)
        
        If Project.Toolbar.Compile Then
            FloodPanel.Value = 0
            FloodPanel.Caption = "Creating Toolbar " + Project.Toolbar.Name + " for " + _
                                IIf(TargetPlatform = tpcBoth, "all supported browsers", IIf(TargetPlatform = tpcIE, "DOM compliant browsers", "Navigator 4"))
            DoEvents
            
            With Project.Toolbar
                If .IsTemplate Then .Condition = "return false;"
                scriptCode = scriptCode + _
                            "tbBorder[" & t & "] = " & .Border & ";" + vbCrLf + vbTab + _
                            "tbSpacing[" & t & "] = " & .Separation & ";" + vbCrLf + vbTab + _
                            "tbStyle[" & t & "] = " & .Style & ";" + vbCrLf + vbTab + _
                            "tbAlignment[" & t & "] = " & .Alignment & ";" + vbCrLf + vbTab + _
                            "tbSpanning[" & t & "] = " & .Spanning & ";" + vbCrLf + vbTab + _
                            "tbFollowHScroll[" & t & "] = " & IIf(.FollowHScroll, "true", "false") & ";" + vbCrLf + vbTab + _
                            "tbFollowVScroll[" & t & "] = " & IIf(.FollowVScroll, "true", "false") & ";" + vbCrLf + vbTab + _
                            "tbMargins[" & t & "] = [" & .OffsetH & ", " & .OffsetV & "];" + vbCrLf + vbTab + _
                            "tbFPos[" & t & "] = [" & .CustX & "," & .CustY & "];" + vbCrLf + vbTab + _
                            "tbVisC[" & t & "] = new Function('" & IIf(LenB(.Condition) <> 0, .Condition, "return true;") & "');" & vbCrLf
                If .IsTemplate Then .Condition = ""
                If LenB(.AttachTo) <> 0 Then
                    scriptCode = scriptCode + _
                            vbTab + "tbAttachTo[" & t & "] = '" & .AttachTo & "|" & .AttachToAlignment & "';" + vbCrLf
                End If
                            
                r = GetTBWidth(t - tbBaseIndex)
                totalWidth = r(1): Wider = r(2)
                r = GetTBHeight(t - tbBaseIndex)
                totalHeight = r(1): Taller = r(2)
                
                scriptCode = scriptCode + vbTab + "tbWidth[" & t & "] = " & IIf(.Style = tscHorizonal, totalWidth, Wider + 2 * (.Border + .ContentsMarginH)) & ";" + vbCrLf
                scriptCode = scriptCode + vbTab + "tbHeight[" & t & "] = " & IIf(.Style = tscHorizonal, Taller + 2 * (.Border + .ContentsMarginV), totalHeight) & ";" + vbCrLf
                
                '******* INTERNET EXPLORER *********************************************
                
                If Project.CompileIECode Or IsForPreviewing Then
                    dw = IIf(.Style = tscHorizonal, IIf(.Spanning = tscAuto, totalWidth & "px", "100%"), (Wider + 2 * (.Border + .ContentsMarginH)) & "px")
                    dH = IIf(.Style = tscVertical, IIf(.Spanning = tscAuto, totalHeight & "px", "100%"), (Taller + 2 * (.Border + .ContentsMarginV)) & "px")
                    ieTable(t) = ieTable(t) + _
                            "<div id=""dmbTBBack" & t & """ style=""position:" + IIf(.Alignment = tbacFree, "relative", "absolute") + "; overflow:hidden; z-index:" & 997 + t & "; visibility:hidden" + _
                                "; width:" & dw & _
                                "; height:" & dH & _
                                RaiusToCSS(.Radius) & _
                                "; " + IIf(Project.Toolbar.Transparency > 0, MozillaAndSafariTrans(Project.Toolbar.Transparency), "") + _
                                IIf(Not IsFlash(.Image), AddConditionalAttribute("; background-image", "url(" + ImgAbsPath + GetFileName(.Image) & ")", False), "") + _
                                AddConditionalAttribute("; background-color", GetRGB(.BackColor, True), False) + _
                                IIf(.BorderColor <> -2 And .Border > 0, "; border:" & .Border & "px " + GetBorderStyleName(.BorderStyle) + " " + GetRGB(.BorderColor, True), "") + _
                                ";""" + IIf(UsesDynAPI, " fh=" & IIf(.Height < 0, Abs(.Height), 0) & " fw=" & IIf(.Width < 0, Abs(.Width), 0) & " cl=" & IIf(.Style = tscHorizonal, 1, 0) & " ld=" & .Separation & " h=" & .ContentsMarginH & " v=" & .ContentsMarginV & " tc=" & .ContentsMarginH & " ss=" & Abs(.JustifyHotSpots) & " lc=" & .ContentsMarginV, "") + ">" + IIf(IsFlash(.Image), CreateFlashCode(.Image, Replace(dw, "px", ""), Replace(dH, "px", ""), GetRGB(.BackColor, True), False), "") + vbCrLf
            
                    ieTable(t) = ieTable(t) + "<div id=""dmbTB" & t & """ style=""position:relative; z-index:" & 998 + t & _
                                "; left:" & (.ContentsMarginH) & "px" + _
                                "; top:" & (.ContentsMarginV) & "px" + _
                                RaiusToCSS(.Radius) & _
                                "; " + IIf(Project.Toolbar.Transparency > 0, MozillaAndSafariTrans(Project.Toolbar.Transparency), "") + _
                                "; " + IIf(Project.Toolbar.DropShadowSize > 0, MozillaAndSafariShadow(Project.Toolbar.DropShadowSize, Project.Toolbar.DropShadowColor), "") + _
                                "; width:" & _
                                IIf(.Style = tscHorizonal, totalWidth, Wider) & "px" & _
                                "; height:" & _
                                IIf(.Style = tscVertical, totalHeight, Taller) & "px" & _
                                ";"">" + vbCrLf
                End If
                
                '******* NAVIGATOR *****************************************************
    
                If Project.CompileNSCode Then
                    dw = IIf(.Style = tscHorizonal, IIf(.Spanning = tscAuto, totalWidth, "100%"), Wider + 2 * (.Border + .ContentsMarginH))
                    dH = IIf(.Style = tscVertical, IIf(.Spanning = tscAuto, totalHeight, "100%"), Taller + 2 * (.Border + .ContentsMarginV))
                    nsTable(t) = nsTable(t) + _
                            "<layer name=dmbTBBack" & t & _
                                " width=" & dw & _
                                " height=" & dH & _
                                    AddConditionalAttribute(" bgColor", GetRGB(.BackColor), True) + _
                                    IIf(Not IsFlash(.Image), AddConditionalAttribute(" background", ImgAbsPath + GetFileName(.Image), True), "") + _
                                " z-index=" & 997 + t & " visibility=hidden>" + _
                                IIf(IsFlash(.Image), CreateFlashCode(.Image, dw, dH, GetRGB(.BackColor), True), "") + vbCrLf
            
                    nsTable(t) = nsTable(t) + "<layer name=dmbTB" & t & _
                                " left=" & (.ContentsMarginH + .Border) & _
                                " top=" & (.ContentsMarginV + .Border) & _
                                " width=" & _
                                IIf(.Style = tscHorizonal, totalWidth, Wider) & _
                                " height=" & _
                                IIf(.Style = tscVertical, totalHeight, Taller) & _
                                " z-index=" & 998 + t & _
                                ">" + vbCrLf
                End If
                
                hsx = 0 '.Border + .ContentsMarginH
                hsy = 0 '.Border + .ContentsMarginV
        
                totalGroupsInTB = UBound(.Groups)
                For i = 1 To UBound(.Groups)
                    g = GetIDByName(.Groups(i))
                    If MenuGrps(g).Compile Then
                        k = k + 1
                        
                        FloodPanel.Value = (i / totalGroupsInTB) * 100
                        
                        hsWidth = GetHotSpotWidth(t - tbBaseIndex, g)
                        hsHeight = GetHotSpotHeight(t - tbBaseIndex, g)
                        
                        'scriptCode = scriptCode + vbTab + "tbHS[" & t & "][" & i & "]=[" & _
                                IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth) & "," & _
                                IIf(.JustifyHotSpots Or .Style = tscHorizonal, Taller, hsHeight) & "];" + vbCrLf
                    
                        mc.Name = MenuGrps(g).Name
                        mc.Caption = MenuGrps(g).Caption
                        mc.LeftImage = MenuGrps(g).tbiLeftImage
                        mc.RightImage = MenuGrps(g).tbiRightImage
                        
                        '******* INTERNET EXPLORER *********************************************
                        
                        If Project.CompileIECode Or IsForPreviewing Then
                            cFX = GetBorderFX(g, True)
                            
                            cStyle = "white-space:nowrap" & _
                                        RaiusToCSS(MenuGrps(g).tbiRadius) & _
                                        IIf(LenB(MenuGrps(g).Caption) <> 0, _
                                            "; font-family:" + GetFontFamily(MenuGrps(g).DefNormalFont.FontName, False) + _
                                            "; font-size:" & MenuGrps(g).DefNormalFont.FontSize & "px" + _
                                            "; font-weight:" + IIf(MenuGrps(g).DefNormalFont.FontBold, "bold", "normal") + _
                                            "; font-style:" + IIf(MenuGrps(g).DefNormalFont.FontItalic, "italic", "normal") + _
                                            "; text-decoration:" + IIf(MenuGrps(g).DefNormalFont.FontUnderline, "underline", "none") + _
                                            "; color:" + GetRGB(MenuGrps(g).nTextColor) + "; ", _
                                            "; overflow:hidden;") + _
                                            GrpShadowToCSS(g, False) + _
                                     cFX(1) + _
                                     IIf(IsFlash(MenuGrps(g).tbiBackImage.NormalImage), "", AddConditionalBackImg(MenuGrps(g).tbiBackImage, False)) + _
                                     AddConditionalAttribute("; background-color", GetRGB(MenuGrps(g).nBackColor, True), False) + ";"
        
                            ieTable(t) = ieTable(t) + "<div id=""N" & k + 1000 & """" & _
                                                " style=""position:absolute" + _
                                                "; top:" & hsy & "px" & _
                                                "; left:" & hsx & "px" & _
                                                "; width:" & IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth) & "px" & _
                                                "; height:" & IIf(.JustifyHotSpots Or .Style = tscHorizonal, Taller, hsHeight) & "px" & _
                                                "; " + cStyle + _
                                                IIf(MenuGrps(g).Disabled, "; cursor:" + GetCursorName(MenuGrps(g).iCursor) + ";", "") + _
                                                """" + _
                                                IIf(UsesDynAPI, " h=" & MenuGrps(g).CmdsMarginX & " v=" & MenuGrps(g).CmdsMarginY, "") & _
                                                IIf(MenuGrps(g).Disabled, ">", _
                                                    GetTBEventCode(g, i, t, False) + ">") + _
                                                CreateTBImageTable(g, False, False, IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth), IIf(.JustifyHotSpots Or .Style = tscHorizonal, Taller, hsHeight)) + "</div>" + vbCrLf
                            
                            If Not MenuGrps(g).Disabled Then
                                cStyle = "white-space:nowrap; font-family:" + GetFontFamily(MenuGrps(g).DefHoverFont.FontName, False) + _
                                         "; font-size:" & MenuGrps(g).DefHoverFont.FontSize & "px" + _
                                         "; font-weight:" + IIf(MenuGrps(g).DefHoverFont.FontBold, "bold", "normal") + _
                                         "; font-style:" + IIf(MenuGrps(g).DefHoverFont.FontItalic, "italic", "normal") + _
                                         "; text-decoration:" + IIf(MenuGrps(g).DefHoverFont.FontUnderline, "underline", "none") + _
                                         "; color:" + GetRGB(MenuGrps(g).hTextColor) + "; " + _
                                         RaiusToCSS(MenuGrps(g).tbiRadius) & _
                                         GrpShadowToCSS(g, True) + cFX(2) + _
                                         IIf(IsFlash(MenuGrps(g).tbiBackImage.HoverImage), "", AddConditionalBackImg(MenuGrps(g).tbiBackImage, True)) + _
                                         AddConditionalAttribute("; background-color", GetRGB(MenuGrps(g).hBackColor, True), False) + _
                                         "; cursor:" + GetCursorName(MenuGrps(g).iCursor) + ";"
                                ieTable(t) = ieTable(t) + "<div id=""O" & k + 1000 & """" & _
                                                    IIf(dispTips And LenB(MenuGrps(g).WinStatus) <> 0, " title=""" + ParseStatusText(g, True) + """ ", "") + _
                                                    " style=""position:absolute" + _
                                                    "; top:" & hsy & "px" & _
                                                    "; left:" & hsx & "px" & _
                                                    "; width:" & IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth) & "px" & _
                                                    "; height:" & IIf(.JustifyHotSpots Or .Style = tscHorizonal, Taller, hsHeight) & "px" & _
                                                    "; " + cStyle + " visibility:hidden;""" + _
                                                    IIf(UsesDynAPI, " h=" & MenuGrps(g).CmdsMarginX & " v=" & MenuGrps(g).CmdsMarginY, "") & _
                                                    IIf(MenuGrps(g).Disabled, ">", _
                                                        GetTBEventCode(g, i, t, True) + ">") + _
                                                    CreateTBImageTable(g, True, False, IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth), IIf(.JustifyHotSpots Or .Style = tscHorizonal, Taller, hsHeight)) + "</div>" + vbCrLf
                            End If
                        End If
                        
                        '******* NAVIGATOR *****************************************************
                        
                        If Project.CompileNSCode Then
                            scHeight = IIf(.JustifyHotSpots Or .Style = tscHorizonal, Taller, hsHeight)
                            CmdLayerName = "MC" & (k + 1000)
                            n = "<ilayer left=" & MenuGrps(g).CmdsMarginX & " top=" & MenuGrps(g).CmdsMarginY & ">" + _
                                    IIf(MenuGrps(g).DefNormalFont.FontBold, "<b>", "") + _
                                    IIf(MenuGrps(g).DefNormalFont.FontItalic, "<i>", "") + _
                                    IIf(MenuGrps(g).DefNormalFont.FontUnderline, "<u>", "") + _
                                    "<font face=" + GetFontFamily(MenuGrps(g).DefNormalFont.FontName, True) + _
                                        " point-size=" & px2pt(MenuGrps(g).DefNormalFont.FontSize) & _
                                        " color=" + GetRGB(MenuGrps(g).nTextColor) + ">" + _
                                        CreateTBImageTable(g, False, True, IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth), IIf(.JustifyHotSpots Or .Style = tscHorizonal, Taller, hsHeight)) + _
                                    "</font>" + _
                                    IIf(MenuGrps(g).DefNormalFont.FontUnderline, "</u>", "") + _
                                    IIf(MenuGrps(g).DefNormalFont.FontItalic, "</i>", "") + _
                                    IIf(MenuGrps(g).DefNormalFont.FontBold, "</b>", "") + _
                                    "</ilayer>"
                            o = "<ilayer left=" & MenuGrps(g).CmdsMarginX & " top=" & MenuGrps(g).CmdsMarginY & ">" + _
                                    IIf(MenuGrps(g).DefHoverFont.FontBold, "<b>", "") + _
                                    IIf(MenuGrps(g).DefHoverFont.FontItalic, "<i>", "") + _
                                    IIf(MenuGrps(g).DefHoverFont.FontUnderline, "<u>", "") + _
                                    "<font face=" + GetFontFamily(MenuGrps(g).DefHoverFont.FontName, True) + _
                                        " point-size=" & px2pt(MenuGrps(g).DefHoverFont.FontSize) & _
                                        " color=" + GetRGB(MenuGrps(g).hTextColor) + ">" + _
                                        CreateTBImageTable(g, True, True, IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth), IIf(.JustifyHotSpots Or .Style = tscHorizonal, Taller, hsHeight)) + _
                                    "</font>" + _
                                    IIf(MenuGrps(g).DefHoverFont.FontUnderline, "</u>", "") + _
                                    IIf(MenuGrps(g).DefHoverFont.FontItalic, "</i>", "") + _
                                    IIf(MenuGrps(g).DefHoverFont.FontBold, "</b>", "") + _
                                    "</ilayer>"
                            
                            nsTable(t) = nsTable(t) + _
                                    "<layer name=" & CmdLayerName & "EH" & CStr(k + 1000) + _
                                        " left=" & hsx & _
                                        " top=" & hsy & _
                                        " width=" & IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth) & _
                                        " height=" & scHeight & _
                                        " z-index=" & 1000 + t & _
                                        IIf(MenuGrps(g).Disabled, ">", _
                                                GetNSTBEventCode(g, i, t)(1) + GetNSTBEventCode(g, i, t)(4) + """>")
                                        nsTable(t) = nsTable(t) + "</layer>"
                                        
                            If LenB(GetNSTBEventCode(g, i, t)(2)) <> 0 Then
                                nsEC = Mid(GetNSTBEventCode(g, i, t)(2), 11)
                                nsEC = Replace(Left(nsEC, Len(nsEC) - 1), """", "\""")
                                nsStyleCode = nsStyleCode + String$(1, vbTab) + "nTCode[" + CStr(k + 1000) + "]=""" + nsEC + """;" + vbCrLf
                            End If
                            
                            nsTable(t) = nsTable(t) + _
                                    "<layer name=" & CmdLayerName & "N" & _
                                        " left=" & hsx & _
                                        " top=" & hsy & _
                                        " width=" & IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth) & _
                                        " height=" & scHeight & _
                                        " z-index=" & 999 + t & _
                                        IIf(IsFlash(MenuGrps(g).tbiBackImage.NormalImage), "", AddConditionalAttribute(" background", ImgAbsPath + GetFileName(MenuGrps(g).tbiBackImage.NormalImage), True, True)) + _
                                        AddConditionalAttribute(" bgColor", GetRGB(MenuGrps(g).nBackColor), True) + _
                                        " visibility=inherit>" + n + _
                                    "</layer>"
                            nsTable(t) = nsTable(t) + _
                                    "<layer name=" & CmdLayerName & "O" & _
                                        " left=" & hsx & _
                                        " top=" & hsy & _
                                        " width=" & IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth) & _
                                        " height=" & scHeight & _
                                        " z-index=" & 999 + t & _
                                        IIf(IsFlash(MenuGrps(g).tbiBackImage.HoverImage), "", AddConditionalAttribute(" background", ImgAbsPath + GetFileName(MenuGrps(g).tbiBackImage.HoverImage), True, True)) + _
                                        AddConditionalAttribute(" bgColor", GetRGB(MenuGrps(g).hBackColor), True) + _
                                        " visibility=hidden>" + o + _
                                    "</layer>"
                        End If
                        Select Case .Style
                            Case tscHorizonal
                                hsx = hsx + IIf(.JustifyHotSpots Or .Style = tscVertical, Wider, hsWidth) + .Separation
                            Case tscVertical
                                hsy = hsy + IIf(.JustifyHotSpots Or .Style = tscHorizonal, Taller, hsHeight) + .Separation
                        End Select
                    End If
                Next i
                
                ' Close the front DIV
                ieTable(t) = ieTable(t) + "</div>"
                nsTable(t) = nsTable(t) + "</layer>"
                
                ' Close the back DIV
                ieTable(t) = ieTable(t) + "</div>"
                nsTable(t) = nsTable(t) + "</layer>"
            End With
        End If
    Next t
    
    Dim IEhasCode As Boolean
    Dim NShasCode As Boolean
    For t = 1 To UBound(ieTable)
        If LenB(ieTable(t)) <> 0 Then IEhasCode = True
        If LenB(nsTable(t)) <> 0 Then NShasCode = True
    Next t
    
SkipBuildTB:
    
    If IsForPreviewing Or TargetPlatform = tpcBoth Then
        If Not IsForPreviewing Then
            If Not Project.CompileIECode Then
                ieStyleCode = ""
                If IEhasCode Then
                    For t = 1 To UBound(ieTable)
                        ieTable(t) = ""
                    Next t
                End If
            End If
            If Not Project.CompileNSCode Then
                nsStyleCode = ""
                If NShasCode Then
                    For t = 1 To UBound(nsTable)
                        nsTable(t) = ""
                    Next t
                End If
            End If
        End If
        
        If IEhasCode Or NShasCode Then
            tbCode(0) = Split(GetTBCodeTemplate, "%%IPOINT%%")(0) + scriptCode
            tbCode(1) = ieStyleCode + nsStyleCode + imgCode + Split(GetTBCodeTemplate, "%%IPOINT%%")(1)
            tbCode(2) = IIf(IEhasCode, GenBrowserTBSaveCode(ieTable, "IE", IsForPreviewing, 1 + tbBaseIndex), "") + _
                        IIf(NShasCode, GenBrowserTBSaveCode(nsTable, "NS", IsForPreviewing, k + tbBaseIndex), "")
        Else
            tbCode(0) = "var lmcHS = null;" + vbCrLf
        End If
        tbCode(0) = tbCode(0) + vbTab + "var tbNum = " & UBound(Project.Toolbars) + tbBaseIndex & ";" + vbCrLf
    Else
        'imgCode = OptimizeImagesCode(imgCode)
        Select Case TargetPlatform
            Case tpcIE
                If Project.CompileIECode Then
                    If IEhasCode Then
                        tbCode(0) = Split(GetTBCodeTemplate, "%%IPOINT%%")(0) + scriptCode
                        tbCode(1) = ieStyleCode + imgCode + Split(GetTBCodeTemplate, "%%IPOINT%%")(1)
                        tbCode(2) = GenBrowserTBSaveCode(ieTable, "IE", IsForPreviewing, 1 + tbBaseIndex)
                    Else
                        tbCode(0) = "var lmcHS = null;" + vbCrLf
                    End If
                    tbCode(0) = tbCode(0) + "var tbNum = " & UBound(Project.Toolbars) + tbBaseIndex & ";" + vbCrLf
                End If
            Case tpcNS
                If Project.CompileNSCode Then
                    If NShasCode Then
                        tbCode(0) = Split(GetTBCodeTemplate, "%%IPOINT%%")(0) + scriptCode
                        tbCode(1) = nsStyleCode + imgCode + Split(GetTBCodeTemplate, "%%IPOINT%%")(1)
                        tbCode(2) = GenBrowserTBSaveCode(nsTable, "NS", IsForPreviewing, 1 + tbBaseIndex)
                    Else
                        tbCode(0) = "var lmcHS = null;" + vbCrLf
                    End If
                    tbCode(0) = tbCode(0) + "var tbNum = " & UBound(Project.Toolbars) + tbBaseIndex & ";" + vbCrLf
                End If
        End Select
    End If
    
    If Project.AutoSelFunction Then
        tbCode(1) = Replace(tbCode(1), "%%AUTOSELP2", "if(mc.disable==true) return false;")
        tbCode(1) = Replace(tbCode(1), "%%AUTOSELP4", "if(isHS>0) return false;")
    Else
        tbCode(1) = Replace(tbCode(1), "%%AUTOSELP2", "")
        tbCode(1) = Replace(tbCode(1), "%%AUTOSELP4", "")
    End If
    tbCode(1) = Replace(tbCode(1), "%%CUSOFFSET", GetCustomOffset("Toolbars"))
    If Project.BlinkEffect > 0 Then
        tbCode(1) = Replace(tbCode(1), "%%BLINKDIS", "if(mibm) return false;")
    Else
        tbCode(1) = Replace(tbCode(1), "%%BLINKDIS", "")
    End If
    
    Dim tmp1 As String
    Dim tmp2 As String
    
    For t = 1 + tbBaseIndex To UBound(Project.Toolbars) + tbBaseIndex
        Project.Toolbar = Project.Toolbars(t - tbBaseIndex)
        With Project.Toolbar
            If .Alignment = tbacAttached And .AttachToAutoResize And .AttachToAlignment = gacBottomLeft Then
                tmp2 = tmp2 + "t==" & t & "&&"
            End If
            
            If .FollowVScroll Then
                If .SmartScrolling = False Then
                    tbCode(1) = Replace(tbCode(1), "if(olt[1]==0 && !tbO[t].svoff", "if(olt[1]==0 && !tbO[t].svoff && t!=" & t)
                    If InStr(1, tbCode(1), "else tbO[t].svoff = -1") = 0 Then
                        tbCode(1) = Replace(tbCode(1), "tbO[t].svoff = getTBPos(t)[1];", "tbO[t].svoff = getTBPos(t)[1]; else tbO[t].svoff = -1;")
                    End If
                Else
                    tbCode(1) = Replace(tbCode(1), "if(olt[1]==0 && !tbO[t].svoff) tbO[t].svoff = getTBPos(t)[1];", "if(olt[1]==0 && !tbO[t].svoff) tbO[t].svoff = parseInt(tbS[t].top) + xOff[1];")
                End If
            End If
        End With
    Next t
    If tmp2 = "" Then
        tbCode(1) = Replace(tbCode(1), "%%IMGAS", "")
    Else
        tmp1 = tmp1 + String(0, vbTab) + "if(imgObj && (%%TBLIST) && tbS[t].visibility=='visible') {" + vbCrLf
        'tmp1 = tmp1 + String(5, vbTab) + "dmbtbF[t].top = parseInt(dmbtbF[t].top) - parseInt(dmbtbB[t].height);" + vbCrLf
        'tmp1 = tmp1 + String(5, vbTab) + "dmbtbB[t].top = parseInt(dmbtbB[t].top) - parseInt(dmbtbB[t].height);" + vbCrLf
        tmp1 = tmp1 + String(5, vbTab) + "imgObj.width = parseInt(tbS[t].width);" + vbCrLf
        tmp1 = tmp1 + String(5, vbTab) + "imgObj.height = parseInt(tbS[t].height);" + vbCrLf
        tmp1 = tmp1 + String(4, vbTab) + "}" + vbCrLf
        
        tbCode(1) = Replace(tbCode(1), "%%IMGAS", Replace(tmp1, "%%TBLIST", Left(tmp2, Len(tmp2) - 2)))
        Project.RemoveImageAutoPosCode = False
    End If
                   
    BuildToolbar = tbCode

End Function

Private Function GetTBCodeTemplate() As String

    Dim tbCode As String
    Dim l() As String
    Dim i As Long
    
    tbCode = LoadFile(AppPath + "rsc\tbcode.dat")
    l = Split(tbCode, vbCrLf)
    For i = 0 To UBound(l)
        l(i) = removeJSComment(l(i))
    Next i
    tbCode = Join(l, vbCrLf)
    If IsMultiProject Then
        tbCode = Mid(tbCode, InStr(tbCode, "%%IPOINT%%"))
    End If
    
    GetTBCodeTemplate = tbCode

End Function

Public Function GenLoaderCode(IsFrames As Boolean, IsForPreviewing As Boolean, Optional ForFile As String, Optional SecondaryProjects As String, Optional XHTMLCompliance As Boolean, Optional YahooSiteBuilderSupport As Boolean) As String

    Dim sStr As String
    Dim fName As String
    Dim relPath2Root As String
    Dim sPrjs() As String
    Dim i As Integer
    Dim tmrFcnStart As String
    Dim tmrFcnEnd As String
    
    Dim rpLib As String
    Dim jPath As String
    
    sPrjs = Split(SecondaryProjects, "|")
    fName = Project.JSFileName + IIf(IsFrames, "_frames", "")
    
    With Project.UserConfigs(Project.DefaultConfig)
        If .Type = ctcCDROM Then
            tmrFcnStart = "function LoadMenus() {if(!rjsPath){window.setTimeout(""LoadMenus()"", 10);return false;}"
            tmrFcnEnd = "return true;}LoadMenus();"
            jPath = "' + rjsPath + '"
        Else
            jPath = JSAbsPath
        End If
        
        If Project.CodeOptimization <> cocDEBUG Then
            sStr = sStr + "<script language=""javascript"" type=""text/javascript"">" + vbCrLf
            sStr = sStr + tmrFcnStart
            sStr = sStr + "var navVer = navigator.appVersion;" + vbCrLf
            sStr = sStr + "if(navVer.substr(0,3) >= 4)" + vbCrLf
            sStr = sStr + "if((navigator.appName==""Netscape"") && (parseInt(navigator.appVersion)==4)) {" + vbCrLf
            
            If Project.CompileNSCode Then
                sStr = sStr + "document.write('<'" + " + 'script language=""javascript"" type=""text/javascript"" src=""" + jPath + "ns" + fName + ".js""><\/script\>');" + vbCrLf
                For i = 1 To UBound(sPrjs)
                    sStr = sStr + "document.write('<'" + " + 'script language=""javascript"" type=""text/javascript"" src=""" + jPath + "ns" + sPrjs(i) + ".js""><\/script\>');" + vbCrLf
                Next i
            Else
                sStr = sStr + ";"
            End If
            
            sStr = sStr + "} else {" + vbCrLf
            
            If Project.CompileIECode Or IsForPreviewing Then
                sStr = sStr + "document.write('<'" + " + 'script language=""javascript"" type=""text/javascript"" src=""" + jPath + "ie" + fName + ".js" + _
                                IIf(Project.UseGZIP And Not IsForPreviewing, "' + (navigator.userAgent.indexOf('Mac')==-1?"".gz"":"""") + '", "") + _
                                """><\/script\>');" + vbCrLf
                For i = 1 To UBound(sPrjs)
                    sStr = sStr + "document.write('<'" + " + 'script language=""javascript"" type=""text/javascript"" src=""" + jPath + "ie" + sPrjs(i) + ".js" + _
                                    IIf(Project.UseGZIP And Not IsForPreviewing, "' + (navigator.userAgent.indexOf('Mac')==-1?"".gz"":"""") + '", "") + _
                                    """><\/script\>');" + vbCrLf
                Next i
            Else
                sStr = sStr + ";"
            End If
            
            sStr = sStr + "}" + tmrFcnEnd
            
            sStr = sStr + "</script>" + vbCrLf
        Else
            sStr = sStr + "<script language=""javascript"" type=""text/javascript"" src=""" + jPath + fName + ".js""></script>" + vbCrLf
            For i = 1 To UBound(sPrjs)
                sStr = sStr + "<script language=""javascript"" type=""text/javascript"" src=""" + jPath + sPrjs(i) + ".js""></script>" + vbCrLf
            Next i
        End If
        
        If .Type = ctcCDROM Then
            relPath2Root = SetSlashDir(GetSmartRelPath(.CompiledPath, .RootWeb), sdFwd)
            
            rpLib = LoadFile(AppPath + "rsc\relpathlib.dat")
            rpLib = Replace(rpLib, "%%RELROOTPATH%%", relPath2Root)
'            If IsFrames Then
'                RemoveFunction rpLib, "FixImages"
'            Else
'                If Project.CodeOptimization = cocDEBUG Then
'                    rpLib = Replace(rpLib, "%%FCN_CALL%%", "f(hStyle);" + IIf(CreateToolbar, "f(hshStyle);", ""))
'                Else
'                    rpLib = Replace(rpLib, "%%FCN_CALL%%", "f(hS);" + IIf(CreateToolbar, "f(hshS);", ""))
'                End If
'            End If
            
            If Project.CodeOptimization = cocDEBUG Then
                sStr = "<script language=""javascript"" type=""text/javascript"">" + vbCrLf + _
                            tmrFcnStart + _
                            "document.write('" + Replace(Replace(Replace(sStr, "<script", "<' + 'script"), vbCrLf, ""), "</script>", "<\/script>") + "');" + _
                            tmrFcnEnd + _
                            "</script>" + vbCrLf
            Else
                rpLib = OptimizeJSCode(rpLib)
            End If
            If LenB(ForFile) = 0 Then ForFile = .HotSpotEditor.HotSpotsFile
            If IsForPreviewing Then
                sStr = vbCrLf + "<div id=""dmbri"" style=""position:absolute;"">" + vbCrLf + _
                        "<img src=""dmb_i.gif"" name=""dmbif"" width=""1"" height=""1"" alt="""">" + vbCrLf + _
                        "<img src=""dmb_m.gif"" name=""dmbjs"" width=""1"" height=""1"" alt="""">" + vbCrLf + _
                        "</div>" + vbCrLf + _
                        "<script language=""javascript"" type=""text/javascript"">" + vbCrLf + _
                        rpLib + vbCrLf + _
                        "</script>" + vbCrLf + _
                        sStr
            Else
                If XHTMLCompliance Then
                    sStr = Replace(sStr, "function LoadMenus", "/* <![CDATA[ */" + "function LoadMenus")
                    sStr = Replace(sStr, "LoadMenus();", "LoadMenus();" + "/* ]]> */")
                    sStr = Replace(sStr, "language=""javascript"" ", "")
                End If
                sStr = vbCrLf + "<div id=""dmbri"" style=""position:absolute;"">" + vbCrLf + _
                        "<img src=""" + SetSlashDir(GetSmartRelPath(GetFilePath(ForFile), .ImagesPath), sdFwd) + "dmb_i.gif" + IIf(Project.LotusDominoSupport, "/$file/dmb_i.gif", "") + """ name=""dmbif"" width=""1"" height=""1"" alt=""""" + IIf(XHTMLCompliance, " /", "") + ">" + vbCrLf + _
                        "<img src=""" + SetSlashDir(GetSmartRelPath(GetFilePath(ForFile), .CompiledPath), sdFwd) + "dmb_m.gif" + IIf(Project.LotusDominoSupport, "/$file/dmb_m.gif", "") + """ name=""dmbjs"" width=""1"" height=""1"" alt=""""" + IIf(XHTMLCompliance, " /", "") + ">" + vbCrLf + _
                        "</div>" + vbCrLf + _
                        "<script" + IIf(XHTMLCompliance, " ", " language=""javascript"" ") + "type=""text/javascript"">" + vbCrLf + _
                        rpLib + vbCrLf + _
                        "</script>" + vbCrLf + _
                        sStr
                If XHTMLCompliance Then
                    sStr = Replace(sStr, " name=""dmbif"" ", " id=""dmbif"" ")
                    sStr = Replace(sStr, " name=""dmbjs"" ", " id=""dmbjs"" ")
                End If
            End If
        End If
    End With
    
    If YahooSiteBuilderSupport Then
        GenLoaderCode = vbCrLf + "  <!--$begin pageHtmlAfter$-->" + LoaderCodeSTART + sStr + GenSEOTweakCode + LoaderCodeEND + " <!--$end pageHtmlAfter$-->" + vbCrLf
    Else
        GenLoaderCode = LoaderCodeSTART + sStr + GenSEOTweakCode + LoaderCodeEND
    End If
    
End Function

Private Function GenSEOTweakCode() As String

    Dim g As Integer
    Dim k As Integer
    Dim html As String
    Dim tidx As Integer
    Dim n As Integer
    
    If Project.SEOTweak = False Then Exit Function
    
    GenSEOTweakCode = Replace(modExport.StdHTML(True), "javascript:void(0);", "#")
    GenSEOTweakCode = Replace(GenSEOTweakCode, "<span>", "")
    GenSEOTweakCode = Replace(GenSEOTweakCode, "</span>", "")
    GenSEOTweakCode = Replace(GenSEOTweakCode, "<a href=""", "<a href=""/")
    GenSEOTweakCode = Replace(GenSEOTweakCode, "<a href=""/http:", "<a href=""http:")
    GenSEOTweakCode = Replace(GenSEOTweakCode, "<a href=""/https:", "<a href=""https:")
    GenSEOTweakCode = Replace(GenSEOTweakCode, "<a href=""/emailto:", "<a href=""emailto:")
    GenSEOTweakCode = Replace(GenSEOTweakCode, "<a href=""/ftp:", "<a href=""ftp:")
    GenSEOTweakCode = Replace(GenSEOTweakCode, "/#", "#")
    
    FloodPanel.Value = 0
    
End Function

Public Function GetEventCode(c As Integer, isOver As Boolean) As String

    Dim cOnMouseOver As String
    Dim cOnClick As String
    Dim cOnDoubleClick As String
    Dim cOnMouseOut As String
    Dim g As Integer
    
    On Error Resume Next
    
    With MenuCmds(c)
    
        g = .Parent
    
        If Not isOver Then
            cOnMouseOver = " OnMouseOver="""
            'If .Sound.OnMouseOver <> "" And .Actions.OnMouseOver.Type <> atcCascade Then
            '    cOnMouseOver = cOnMouseOver + "cFrame.SelSndId=" & GetSoundID(.Sound.OnMouseOver) & ";"
            'End If
            cOnMouseOver = cOnMouseOver + "cFrame.HoverSel(0,this);"
            Select Case .Actions.OnMouseOver.Type
                Case atcCascade
                    'If Project.UnfoldingSound.OnMouseOver <> "" Then
                    '    cOnMouseOver = cOnMouseOver + "cFrame.SelSndId=" & GetSoundID(Project.UnfoldingSound.OnMouseOver) & ";"
                    'End If
                    cOnMouseOver = cOnMouseOver + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnMouseOver.TargetMenu).Name + "', 0, 0, true, '', " & .Actions.OnMouseOver.TargetMenuAlignment & ");"
                Case atcURL
                    #If DEMO = 0 Then
                    If Not IsDEMO Then
                        If LCase(Left(.Actions.OnMouseOver.url, 11)) = "javascript:" Then
                            cOnMouseOver = cOnMouseOver + Mid(.Actions.OnMouseOver.url, 12) + ";"
                        Else
                            cOnMouseOver = cOnMouseOver + "cFrame.execURL('" + FixURL(.Actions.OnMouseOver.url) + "', '" + GenFramesObject(.Actions.OnMouseOver.TargetFrame) + "');"
                        End If
                    End If
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnMouseOver = cOnMouseOver + GenNewWindowCode(.Actions.OnMouseOver)
                    #End If
            End Select
            If LenB(.WinStatus) <> 0 Then cOnMouseOver = cOnMouseOver + "window.status=cFrame._purl('" + ParseStatusText(c, False) + "');"
            cOnMouseOver = cOnMouseOver + """"
        End If
        
        If .Actions.OnClick.Type <> atcNone And isOver Then
            cOnClick = " OnClick="""
            'If .Sound.OnClick <> "" And .Actions.OnClick.Type <> atcCascade Then
            '    cOnClick = cOnClick + "cFrame.SelSndId=" & GetSoundID(.Sound.OnClick) & ";"
            'End If
            Select Case .Actions.OnClick.Type
                Case atcCascade
                    'If Project.UnfoldingSound.OnClick <> "" Then
                    '    cOnClick = cOnClick + "cFrame.SelSndId=" & GetSoundID(Project.UnfoldingSound.OnMouseOver) & ";"
                    'End If
                    cOnClick = cOnClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnClick.TargetMenu).Name + "', 0, 0, true, '', " & .Actions.OnClick.TargetMenuAlignment & ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + "cFrame.execURL('" + FixURL(.Actions.OnClick.url) + "', '" + GenFramesObject(.Actions.OnClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + GenNewWindowCode(.Actions.OnClick)
                    #End If
            End Select
            cOnClick = cOnClick + """"
        End If
        
        If .Actions.OnDoubleClick.Type <> atcNone And isOver Then
            cOnDoubleClick = " OnDblClick="""
            Select Case .Actions.OnDoubleClick.Type
                Case atcCascade
                    'If Project.UnfoldingSound.OnClick <> "" Then
                    '    cOnDoubleClick = cOnDoubleClick + "cFrame.SelSndId=" & GetSoundID(Project.UnfoldingSound.OnMouseOver) & ";"
                    'End If
                    cOnDoubleClick = cOnDoubleClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnDoubleClick.TargetMenu).Name + "', 0, 0, true, '', " & .Actions.OnDoubleClick.TargetMenuAlignment & ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + "cFrame.execURL('" + FixURL(.Actions.OnDoubleClick.url) + "', '" + GenFramesObject(.Actions.OnDoubleClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + GenNewWindowCode(.Actions.OnDoubleClick)
                    #End If
            End Select
            cOnDoubleClick = cOnDoubleClick + """"
        End If
    End With
    
    GetEventCode = cOnMouseOver + cOnClick + cOnDoubleClick
    
End Function

Private Function GetNSEventCode(c As Integer) As String()

    Dim cOnMouseOver As String
    Dim cOnClick As String
    Dim cOnDoubleClick As String
    Dim cEventCode(1 To 3) As String
    Dim g As Integer
    
    On Error Resume Next
    
    With MenuCmds(c)
    
        g = .Parent
    
        cOnMouseOver = " OnMouseOver="""
        'If .Sound.OnMouseOver <> "" And .Actions.OnMouseOver.Type <> atcCascade Then
        '    cOnMouseOver = cOnMouseOver + "cFrame.SelSndId=" & GetSoundID(.Sound.OnMouseOver) & ";"
        'End If
        cOnMouseOver = cOnMouseOver + "cFrame.NSHoverSel(0,this);status='" + ParseStatusText(c, False, True) + "';"
        Select Case .Actions.OnMouseOver.Type
            Case atcCascade
                'If Project.UnfoldingSound.OnMouseOver <> "" Then
                '    cOnMouseOver = cOnMouseOver + "cFrame.SelSndId=" & GetSoundID(Project.UnfoldingSound.OnMouseOver) & ";"
                'End If
                cOnMouseOver = cOnMouseOver + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnMouseOver.TargetMenu).Name + "', 0, 0, true, '', " & .Actions.OnMouseOver.TargetMenuAlignment & ");"
            Case atcURL
                #If DEMO = 0 Then
                cOnMouseOver = cOnMouseOver + "cFrame.execURL('" + FixURL(.Actions.OnMouseOver.url) + "', '" + GenFramesObject(.Actions.OnMouseOver.TargetFrame) + "');"
                #End If
            Case atcNewWindow
                #If DEMO = 0 Then
                cOnMouseOver = cOnMouseOver + GenNewWindowCode(.Actions.OnMouseOver)
                #End If
        End Select
        cOnMouseOver = cOnMouseOver + """"
        
        If .Actions.OnClick.Type <> atcNone Then
            cOnClick = " OnClick="""
            'If .Sound.OnClick <> "" And .Actions.OnClick.Type <> atcCascade Then
            '    cOnClick = cOnClick + "cFrame.SelSndId=" & GetSoundID(.Sound.OnClick) & ";"
            'End If
            Select Case .Actions.OnClick.Type
                Case atcCascade
                    'If Project.UnfoldingSound.OnMouseOver <> "" Then
                    '    cOnClick = cOnClick + "cFrame.SelSndId=" & GetSoundID(Project.UnfoldingSound.OnMouseOver) & ";"
                    'End If
                    cOnClick = cOnClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnClick.TargetMenu).Name + "', 0, 0, true, '', " & .Actions.OnMouseOver.TargetMenuAlignment & ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + "cFrame.execURL('" + FixURL(.Actions.OnClick.url) + "', '" + GenFramesObject(.Actions.OnClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + GenNewWindowCode(.Actions.OnClick)
                    #End If
            End Select
            cOnClick = cOnClick + """"
        End If
        
        If .Actions.OnDoubleClick.Type <> atcNone Then
            cOnDoubleClick = " OnDblClick="""
            Select Case .Actions.OnDoubleClick.Type
                Case atcCascade
                    'If Project.UnfoldingSound.OnMouseOver <> "" Then
                    '    cOnDoubleClick = cOnDoubleClick + "cFrame.SelSndId=" & GetSoundID(Project.UnfoldingSound.OnMouseOver) & ";"
                    'End If
                    cOnDoubleClick = cOnDoubleClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnDoubleClick.TargetMenu).Name + "', 0, 0, true, '', " & .Actions.OnMouseOver.TargetMenuAlignment & ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + "cFrame.execURL('" + FixURL(.Actions.OnDoubleClick.url) + "', '" + GenFramesObject(.Actions.OnDoubleClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + GenNewWindowCode(.Actions.OnDoubleClick)
                    #End If
            End Select
            cOnDoubleClick = cOnDoubleClick + """"
        End If
    End With
    
    cEventCode(1) = cOnMouseOver
    cEventCode(2) = cOnClick
    cEventCode(3) = cOnDoubleClick
    
    GetNSEventCode = cEventCode
    
End Function

Private Function GetTBEventCode(g As Integer, i As Integer, t As Integer, isOver As Boolean) As String

    Dim cOnMouseOver As String
    Dim cOnClick As String
    Dim cOnDoubleClick As String
    Dim cOnMouseOut As String
    Dim rg As Integer
    Dim DispSubMenu As Boolean
    
    rg = GetRealSubGroup(g)
    With MenuGrps(g)
        DispSubMenu = (.Actions.OnMouseOver.Type = atcCascade) Or _
                        (.Actions.OnClick.Type = atcCascade) Or _
                        (.Actions.OnDoubleClick.Type = atcCascade)
    
        cOnMouseOver = " OnMouseOver="""
        If Not isOver Then
            cOnMouseOver = cOnMouseOver + "hsHoverSel(0,this);"
            'If Not DispSubMenu Then cOnMouseOver = cOnMouseOver + "HideAll(1);"
        ElseIf Not DispSubMenu Then
            'cOnMouseOver = cOnMouseOver + "mhdHnd = window.setTimeout('HideAll(1)', scDelay);"
        End If
        
        Select Case .Actions.OnMouseOver.Type
            Case atcCascade
                cOnMouseOver = cOnMouseOver + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnMouseOver.TargetMenu).Name + "'," & i & "," & MenuGrps(g).Alignment & ",false,'',-" & t & ");"
            Case atcURL
                #If DEMO = 0 Then
                If LCase(Left(.Actions.OnMouseOver.url, 11)) = "javascript:" Then
                    cOnMouseOver = cOnMouseOver + Mid(.Actions.OnMouseOver.url, 12) + ";"
                Else
                    cOnMouseOver = cOnMouseOver + "cFrame.execURL('" + FixURL(.Actions.OnMouseOver.url) + "', '" + GenFramesObject(.Actions.OnMouseOver.TargetFrame) + "');"
                End If
                #End If
            Case atcNewWindow
                #If DEMO = 0 Then
                cOnMouseOver = cOnMouseOver + GenNewWindowCode(.Actions.OnMouseOver)
                #End If
        End Select
        If LenB(.WinStatus) <> 0 Then cOnMouseOver = cOnMouseOver + "window.status=cFrame._purl('" + ParseStatusText(g, True) + "');"
        cOnMouseOver = cOnMouseOver + """"
        
        If cOnMouseOver = " OnMouseOver=""""" Then
            cOnMouseOver = ""
        End If
        If isOver Then
'            cOnMouseOut = " OnMouseOut=""IsOverHS=false;"""
'            If Not DispSubMenu Then
'                cOnMouseOut = " OnMouseOut=""if(lmcHS)hsHoverSel(1);"""
'            Else
'                If .Actions.OnClick.Type = atcCascade Then
'                    cOnMouseOut = " OnMouseOut=""if(!InMenu())tHideAll();else cFrame.IsOverHS=false;"""
'                Else
'                    cOnMouseOut = " OnMouseOut=""tHideAll();"""
'                End If
'            End If
        End If
        
        If .Actions.OnClick.Type <> atcNone Then
            cOnClick = " OnClick="""
            Select Case .Actions.OnClick.Type
                Case atcCascade
                    cOnClick = cOnClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnClick.TargetMenu).Name + "'," & i & "," & MenuGrps(g).Alignment & ",false,'',-" & t & ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + "cFrame.execURL('" + FixURL(.Actions.OnClick.url) + "', '" + GenFramesObject(.Actions.OnClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + GenNewWindowCode(.Actions.OnClick)
                    #End If
            End Select
            cOnClick = cOnClick + """"
        End If
        
        If .Actions.OnDoubleClick.Type <> atcNone Then
            cOnDoubleClick = " OnDblClick="""
            Select Case .Actions.OnDoubleClick.Type
                Case atcCascade
                    cOnDoubleClick = cOnDoubleClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnMouseOver.TargetMenu).Name + "'," & i & "," & MenuGrps(g).Alignment & ",false,'',-" & t & ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + "cFrame.execURL('" + FixURL(.Actions.OnDoubleClick.url) + "', '" + GenFramesObject(.Actions.OnDoubleClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + GenNewWindowCode(.Actions.OnDoubleClick)
                    #End If
            End Select
            cOnDoubleClick = cOnDoubleClick + """"
        End If
    End With
    
    GetTBEventCode = cOnMouseOver + cOnClick + cOnDoubleClick + cOnMouseOut
    
End Function

Public Function GetNSTBEventCode(g As Integer, i As Integer, t As Integer) As String()

    Dim cOnMouseOver As String
    Dim cOnClick As String
    Dim cOnDoubleClick As String
    Dim cOnMouseOut As String
    Dim eCode(1 To 4) As String
    Dim rg As Integer
    
    rg = GetRealSubGroup(g)
    With MenuGrps(g)
        cOnMouseOver = " OnMouseOver="""
        'If .Actions.OnMouseOver.Type <> atcCascade Then cOnMouseOver = cOnMouseOver + "if(!HTHnd[nOM])cFrame.IsOverHS=true;"
        cOnMouseOver = cOnMouseOver + "hsNSHoverSel(0,this);"
        If LenB(.WinStatus) <> 0 Then cOnMouseOver = cOnMouseOver + "window.status=cFrame._purl('" + ParseStatusText(g, True, True) + "');"
        Select Case .Actions.OnMouseOver.Type
            Case atcCascade
                cOnMouseOver = cOnMouseOver + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnMouseOver.TargetMenu).Name + "'," & i & "," & MenuGrps(g).Alignment & ",false,'',-" & t & ");"
            Case atcURL
                #If DEMO = 0 Then
                cOnMouseOver = cOnMouseOver + "cFrame.execURL('" + FixURL(.Actions.OnMouseOver.url) + "', '" + GenFramesObject(.Actions.OnMouseOver.TargetFrame) + "');"
                #End If
            Case atcNewWindow
                #If DEMO = 0 Then
                cOnMouseOver = cOnMouseOver + GenNewWindowCode(.Actions.OnMouseOver)
                #End If
        End Select
        cOnMouseOver = cOnMouseOver + """"
        
        'If .Actions.OnMouseOver.Type <> atcCascade And .Actions.OnClick.Type <> atcCascade And .Actions.OnDoubleClick.Type <> atcCascade Then
        '    cOnMouseOut = " OnMouseOut=""if(lmcHS)hsNSHoverSel(1);"""
        'Else
        '    cOnMouseOut = " OnMouseOut=""tHideAll();"""
        'End If
        
        If .Actions.OnClick.Type <> atcNone Then
            cOnClick = " OnClick="""
            Select Case .Actions.OnClick.Type
                Case atcCascade
                    cOnClick = cOnClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnClick.TargetMenu).Name + "'," & i & "," & MenuGrps(g).Alignment & ",false,'',-" & t & ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + "cFrame.execURL('" + FixURL(.Actions.OnClick.url) + "', '" + GenFramesObject(.Actions.OnClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + GenNewWindowCode(.Actions.OnClick)
                    #End If
            End Select
            cOnClick = cOnClick + """"
        End If
        
        If .Actions.OnDoubleClick.Type <> atcNone Then
            cOnDoubleClick = " OnDblClick="""
            Select Case .Actions.OnDoubleClick.Type
                Case atcCascade
                    cOnDoubleClick = cOnDoubleClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnMouseOver.TargetMenu).Name + "'," & i & "," & MenuGrps(g).Alignment & ",false,'',-" & t & ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + "cFrame.execURL('" + FixURL(.Actions.OnDoubleClick.url) + "', '" + GenFramesObject(.Actions.OnDoubleClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + GenNewWindowCode(.Actions.OnDoubleClick)
                    #End If
            End Select
            cOnDoubleClick = cOnDoubleClick + """"
        End If
    End With
    
    eCode(1) = cOnMouseOver
    eCode(2) = cOnClick
    eCode(3) = cOnDoubleClick
    eCode(4) = cOnMouseOut
    
    GetNSTBEventCode = eCode
    
End Function

Private Sub AddExpItem(sStr As String, Platform As TargetPlatformConstants, Optional chkSize As Boolean = False)

    If chkSize Then
        If Len(sStr) < 5 Then Exit Sub
    End If

    ReDim Preserve ExpandableItems(UBound(ExpandableItems) + 1)
    
    With ExpandableItems(UBound(ExpandableItems))
        .sStr = sStr
        .Platform = Platform
    End With

End Sub

Private Sub Class_Initialize()

    DoUNICODE = (GetSetting("DMB", "Preferences", "DoUNICODE", 1) = 1)
    cChar = Chr(1)

End Sub

Public Function GetGroupEventCode(g As Integer, ByVal HSImgName As String, IsForPreviewing As Boolean) As String

    Dim cOnMouseOver As String
    Dim cOnClick As String
    Dim cOnDoubleClick As String
    Dim XYPos As String
    Dim ExtraParams As String
    
    Dim i As Integer
    Dim j As Integer
    
    cOnMouseOver = "OnMouseOver="""
    cOnClick = "OnClick="""
    cOnDoubleClick = "OnDblClick="""
    
    On Error Resume Next
    
    XYPos = GetXYPos(g, IsForPreviewing)
    If LenB(HSImgName) = 0 Or IsForPreviewing Then
        ExtraParams = ""
    Else
        HSImgName = HSImgName + "|3"
        If UBound(MenuGrps) > 1 And Project.UserConfigs(Project.DefaultConfig).Frames.UseFrames Then
            HSImgName = Left(HSImgName, Len(HSImgName) - 2)
            For i = 1 To UBound(MenuGrps)
                If MenuGrps(i).x <> MenuGrps(i).y And MenuGrps(i).Compile Then
                    For j = i + 1 To UBound(MenuGrps)
                        If MenuGrps(j).x <> MenuGrps(j).y And MenuGrps(j).Compile Then
                            If MenuGrps(i).x = MenuGrps(j).x Then
                                HSImgName = HSImgName + "|1"
                            Else
                                HSImgName = HSImgName + "|2"
                            End If
                            Exit For
                        End If
                    Next j
                    Exit For
                End If
            Next i
        End If
        ExtraParams = ",'" + HSImgName + "'," & MenuGrps(g).Alignment
    End If
    
    With MenuGrps(g)
        Select Case .Actions.OnMouseOver.Type
            Case atcCascade
                cOnMouseOver = cOnMouseOver + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnMouseOver.TargetMenu).Name + "', " + XYPos + ", false" + ExtraParams + ");"
            Case atcURL
                #If DEMO = 0 Then
                cOnMouseOver = cOnMouseOver + "cFrame.execURL('" + FixURL(.Actions.OnMouseOver.url) + "', '" + GenFramesObject(.Actions.OnMouseOver.TargetFrame) + "');"
                #End If
            Case atcNewWindow
                #If DEMO = 0 Then
                cOnMouseOver = cOnMouseOver + GenNewWindowCode(.Actions.OnMouseOver)
                #End If
        End Select
        cOnMouseOver = cOnMouseOver + """"
        
        If .Actions.OnClick.Type <> atcNone Then
            cOnClick = " OnClick="""
            Select Case .Actions.OnClick.Type
                Case atcCascade
                    cOnClick = cOnClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnClick.TargetMenu).Name + "', " + XYPos + ", false" + ExtraParams + ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + "cFrame.execURL('" + FixURL(.Actions.OnClick.url) + "', '" + GenFramesObject(.Actions.OnClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnClick = cOnClick + GenNewWindowCode(.Actions.OnClick)
                    #End If
            End Select
            cOnClick = cOnClick + """"
        End If
        
        If .Actions.OnDoubleClick.Type <> atcNone Then
            cOnDoubleClick = " OnDblClick="""
            Select Case .Actions.OnDoubleClick.Type
                Case atcCascade
                    cOnDoubleClick = cOnDoubleClick + "cFrame.ShowMenu('" + MenuGrps(.Actions.OnDoubleClick.TargetMenu).Name + "', " + XYPos + ", false" + ExtraParams + ");"
                Case atcURL
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + "cFrame.execURL('" + FixURL(.Actions.OnDoubleClick.url) + "', '" + GenFramesObject(.Actions.OnDoubleClick.TargetFrame) + "');"
                    #End If
                Case atcNewWindow
                    #If DEMO = 0 Then
                    cOnDoubleClick = cOnDoubleClick + GenNewWindowCode(.Actions.OnDoubleClick)
                    #End If
            End Select
            cOnDoubleClick = cOnDoubleClick + """"
        End If
    End With
    
    If cOnMouseOver = "OnMouseOver=""" Then cOnMouseOver = ""
    If cOnClick = "OnClick=""" Then cOnClick = ""
    If cOnDoubleClick = "OnDblClick=""" Then cOnDoubleClick = ""
    
    GetGroupEventCode = cOnMouseOver + cOnClick + cOnDoubleClick + " OnMouseOut=""cFrame.tHideAll();"""
    
End Function

Public Function DecodeNWP(url As String, par As String) As String

    Dim dec As String
    Dim t As String
    Dim l As String
    Dim w As String
    Dim h As String
    
    On Error Resume Next
    
    If LenB(par) = 0 Then
        par = nwdPar
    End If
    
    l = IIf(LenB(GetParam(par, 2)) <> 0, GetParam(par, 2), "")
    t = IIf(LenB(GetParam(par, 3)) <> 0, GetParam(par, 3), "")
    w = IIf(LenB(GetParam(par, 4)) <> 0, GetParam(par, 4), "")
    h = IIf(LenB(GetParam(par, 5)) <> 0, GetParam(par, 5), "")
    
    If InStr(w, "%") Then w = Val(w) / 100 & "*screen.width"
    If InStr(h, "%") Then h = Val(h) / 100 & "*screen.height"
    If l = "*" Then
        l = "(screen.width - " & Val(w) & ")/2"
        t = "(screen.height - " & Val(h) & ")/2"
    Else
        If InStr(l, "%") Then l = Val(l) / 100 & "*screen.width"
        If InStr(t, "%") Then t = Val(t) / 100 & "*screen.height"
    End If
    
    If LenB(l) <> 0 Then l = IIf(CStr(CInt(Val(l))) = l, "left=" & l & ",", "left='+" + l + "+',")
    If LenB(t) <> 0 Then t = IIf(CStr(CInt(Val(t))) = t, "top=" & t & ",", "top='+" + t + "+',")
    If LenB(w) <> 0 Then w = IIf(CStr(CInt(Val(w))) = w, "width=" & w & ",", "width='+" + w + "+',")
    If LenB(h) <> 0 Then h = IIf(CStr(CInt(Val(h))) = h, "height=" & h & ",", "height='+" + h + "+',")
    
    dec = "cFrame.HideAll();cFrame.dmbNW=window.open(cFrame._purl('" + url + "'),'" + GetParam(par, 1) + "','" + _
          l + _
          t + _
          w + _
          h + _
          "directories=" + GetParam(par, 6) + "," + _
          "channelmode=" + GetParam(par, 7) + "," + _
          "toolbar=" + GetParam(par, 8) + "," + _
          "fullscreen=" + GetParam(par, 9) + "," + _
          "location=" + GetParam(par, 10) + "," + _
          "menubar=" + GetParam(par, 11) + "," + _
          "resizable=" + GetParam(par, 12) + "," + _
          "scrollbars=" + GetParam(par, 13) + "," + _
          "status=" + GetParam(par, 14) + "," + _
          "titlebar=" + GetParam(par, 15) + _
          "');if(cFrame.IsFrames&&IE&&!SM)mFrame.location.reload();if(cFrame.dmbNW)cFrame.dmbNW.focus();"
          
    DecodeNWP = Replace(dec, "fullscreen=0", "")

End Function

Public Function GenNewWindowCode(TheAction As Action) As String

    GenNewWindowCode = DecodeNWP(FixURL(TheAction.url), TheAction.WindowOpenParams)

End Function

Private Function GenComment(pTitle As String, pValue As String) As String

    GenComment = "// " & pTitle & " " & pValue & Space$(Abs(44 - (Len(pTitle) + Len(pValue)))) + "//" + vbCrLf

End Function

Public Function IsDEMO() As Boolean

    #If DEMO = 1 Then
        IsDEMO = True
    #Else
        Dim c As String

        Dim User As String
        Dim Company As String

        On Error Resume Next

        c = HEX2Str(Inflate(HEX2Str(GetSetting("DMB", "RegInfo", "CacheData"))))

        User = GetSetting("DMB", "RegInfo", "User")
        Company = GetSetting("DMB", "RegInfo", "Company")

        IsDEMO = Not ((User = Split(c, "|")(0)) And (Company = Split(c, "|")(1)))

        If IsDEMO Then SaveSetting "DMB", "RegInfo", "CacheSig02", FileLen(App.Path + "\" + App.EXEName + ".dll")
        'IsDEMO = True
    #End If

End Function

Private Sub SetupExpandableKeywords(Platform As TargetPlatformConstants)

    Dim FontFamily As String
    Dim i As Integer
    Dim j As Integer
    Dim tmpStr As String

    Erase ExpandableItems: ReDim ExpandableItems(0)

    If Platform = tpcIE Then
        For i = 1 To UBound(MenuGrps)
            If LenB(GetBorderFX(i, True)(1)) <> 0 Then AddExpItem Replace(GetBorderFX(i, True)(1), ": ", ":"), tpcIE
            If LenB(GetBorderFX(i, True)(2)) <> 0 Then AddExpItem Replace(GetBorderFX(i, True)(2), ": ", ":"), tpcIE
        Next i
        
        If UBound(MenuCmds) > 0 Then
            AddExpItem GetRGB(MenuCmds(1).nTextColor, True), tpcIE
            AddExpItem GetRGB(MenuCmds(1).nBackColor, True), tpcIE
            AddExpItem GetRGB(MenuCmds(1).hTextColor, True), tpcIE
            AddExpItem GetRGB(MenuCmds(1).hBackColor, True), tpcIE
        End If
        
        AddExpItem "\"" style=\""position:absolute;top:", tpcIE
        AddExpItem "px;left:", tpcIE
        AddExpItem "px;width:", tpcIE
        AddExpItem "px;height:", tpcIE
        
        AddExpItem "px;visibility:hidden;", tpcIE
        AddExpItem "progid:DXImageTransform.Microsoft.", tpcIE
        AddExpItem "-moz-opacity:", tpcIE
        AddExpItem "(direction=", tpcIE
        AddExpItem ",color=", tpcIE
        AddExpItem ",strength=", tpcIE
        AddExpItem "opacity=", tpcIE
        AddExpItem ")\""><div id=\""", tpcIE
        
        AddExpItem ";background-color:", tpcIE
        AddExpItem ";background-image:url(", tpcIE
        AddExpItem ".gif);border-left:", tpcIE
        AddExpItem ".jpg);border-left:", tpcIE
        AddExpItem ".png);border-left:", tpcIE
        
        i = -1
        Do
            i = i + 1
            tmpStr = GetBorderStyleName(i)
            AddExpItem "px " + tmpStr + ";border-right:", tpcIE
            AddExpItem "px " + tmpStr + ";border-top:", tpcIE
            AddExpItem "px " + tmpStr + ";border-bottom:", tpcIE
            AddExpItem "px " + tmpStr + ";\""><div style=\""position:absolute;top:", tpcIE
        Loop Until tmpStr = "solid" And i > 0
        
        AddExpItem "px;white-space:nowrap;font-family:", tpcIE
        AddExpItem ";font-size:", tpcIE
        
        AddExpItem "px;font-weight:normal;font-style:normal;text-decoration:none;color:", tpcIE
        AddExpItem "px;font-weight:bold;font-style:normal;text-decoration:none;color:", tpcIE
        AddExpItem "px;font-weight:normal;font-style:italic;text-decoration:none;color:", tpcIE
        AddExpItem "px;font-weight:bold;font-style:italic;text-decoration:none;color:", tpcIE
        AddExpItem "px;font-weight:normal;font-style:normal;text-decoration:underline;color:", tpcIE
        AddExpItem "px;font-weight:bold;font-style:normal;text-decoration:underline;color:", tpcIE
        AddExpItem "px;font-weight:normal;font-style:italic;text-decoration:underline;color:", tpcIE
        AddExpItem "px;font-weight:bold;font-style:italic;text-decoration:underline;color:", tpcIE
        
        AddExpItem ";\"" id=""N", tpcIE
        AddExpItem " OnMouseOver=\""cFrame.HoverSel(0,this);\""><img style=\""position:absolute;top:", tpcIE
        
        AddExpItem "px;\"" src=\""", tpcIE
        AddExpItem ".gif\"" width=", tpcIE
        AddExpItem ".jpg\"" width=", tpcIE
        AddExpItem ".png\"" width=", tpcIE
        AddExpItem " height=", tpcIE
        AddExpItem "><div style=\""position:absolute;top:", tpcIE
        AddExpItem "px;\"" align=""left"">", tpcIE
        AddExpItem "px;\"" align=""right"">", tpcIE
        AddExpItem "px;\"" align=""center"">", tpcIE
        
        AddExpItem ";border:", tpcIE
        
        Dim c As tCursor
        c.cType = 0
        Do
            c.cType = c.cType + 1
            tmpStr = GetCursorName(c)
            If tmpStr = "" Then Exit Do
            AddExpItem ";cursor:" + tmpStr + ";visibility:hidden;\""", tpcIE
        Loop
        
        AddExpItem "</div></div", tpcIE
        
        AddExpItem " OnClick=\""cFrame.execURL(\'%%REP%%", tpcIE
        AddExpItem " OnClick=\""cFrame.execURL(\'", tpcIE
        AddExpItem "\',\'_self\');\""><img style=\""position:absolute;top:", tpcIE
        AddExpItem "><div id=\""", tpcIE
        
        AddExpItem " OnMouseOver=\""cFrame.HoverSel(0,this);cFrame.ShowMenu(\'", tpcIE
        AddExpItem "\',0,0,true,\'\',6);\""><img style=\""position:absolute;top:", tpcIE
        AddExpItem "</div><img style=\""position:absolute;top:", tpcIE
        AddExpItem "px solid ", tpcIE
        AddExpItem "><img style=\""position:absolute;top:", tpcIE
        AddExpItem " style=\""position:absolute;overflow:hidden;z-index:", tpcIE
        AddExpItem ";visibility:hidden;width:", tpcIE
        AddExpItem ";filter:", tpcIE
        AddExpItem ";\""></div><div id=""dmbTB", tpcIE
        AddExpItem " style=\""position:absolute;z-index:9", tpcIE
        AddExpItem "><div id=""N1", tpcIE
        AddExpItem " style=\""position:absolute;", tpcIE
        AddExpItem "\"" OnMouseOver=\""hsHoverSel(0,this);cFrame.ShowMenu(\'", tpcIE
        AddExpItem ",0,false,\'\',-1);\""", tpcIE
        AddExpItem "</div></div><div id=""O1", tpcIE
        AddExpItem "OnMouseOver=\""cFrame.ShowMenu(\'", tpcIE
        AddExpItem "</div></div", tpcIE
        AddExpItem "<div id=""dmbTBBack", tpcIE

        AddExpItem ";\"" OnMouseOver=\""hsHoverSel(0,this);window.status=cFrame._purl(\'", tpcIE
        AddExpItem "\');\""", tpcIE
        AddExpItem "></div><div id=""O1", tpcIE
        AddExpItem "OnMouseOver=\""window.status=cFrame._purl(\'", tpcIE
        AddExpItem "></div", tpcIE
        AddExpItem ",0,false,\'\',-4);window.status=cFrame._purl(\'", tpcIE
        AddExpItem " OnMouseOver=\""cFrame.HoverSel(0,this);window.status=cFrame._purl(\'", tpcIE
        AddExpItem "px;white-space:nowrap;overflow:hidden", tpcIE
        AddExpItem "\',0,0,true,\'\',6);window.status=cFrame._purl(\'", tpcIE
        
        AddExpItem "overflow:hidden;z-index:9", tpcIE
        AddExpItem ";\"" OnMouseOver=\""hsHoverSel(0,this);\""", tpcIE
        AddExpItem " OnMouseOver=\""cFrame.HoverSel(0,this);\""", tpcIE
        AddExpItem "border-left:", tpcIE
        AddExpItem ";border-top:", tpcIE
        AddExpItem ";border-right:", tpcIE
        AddExpItem ";border-bottom:", tpcIE
        
        i = -1
        Do
            i = i + 1
            tmpStr = GetBorderStyleName(i)
            AddExpItem "px " + tmpStr + " ", tpcIE
        Loop Until tmpStr = "solid" And i > 0
        
        AddExpItem " id=""O", tpcIE
        
        If Project.CodeOptimization = cocAggressive Then
            For i = 1 To UBound(MenuCmds)
                With MenuCmds(i)
                    If .LeftImage.NormalImage <> "" Then AddExpItem GetFileName(.LeftImage.NormalImage, True), tpcIE, True
                    If .LeftImage.HoverImage <> "" Then AddExpItem GetFileName(.LeftImage.HoverImage, True), tpcIE, True
                    If .BackImage.NormalImage <> "" Then AddExpItem GetFileName(.BackImage.NormalImage, True), tpcIE, True
                    If .BackImage.HoverImage <> "" Then AddExpItem GetFileName(.BackImage.HoverImage, True), tpcIE, True
                    If .RightImage.NormalImage <> "" Then AddExpItem GetFileName(.RightImage.NormalImage, True), tpcIE, True
                    If .RightImage.HoverImage <> "" Then AddExpItem GetFileName(.RightImage.HoverImage, True), tpcIE, True
                End With
            Next i
        End If
        
        AddExpItem " title=\""", tpcIE
        AddExpItem "window.status=cFrame._purl(\'", tpcIE
        AddExpItem ",0,false,\'\',-1);", tpcIE
        
        AddExpItem "\',\'frames[top].frames[", tpcIE
        AddExpItem " OnClick=\""cFrame.HideAll();if(cFrame.dmbNW)cFrame.dmbNW=window.open(cFrame._purl(\'", tpcIE
        AddExpItem "\');if(cFrame.IsFrames&&IE&&!SM)mFrame.location.reload();if(cFrame.dmbNW)cFrame.dmbNW.focus();\""", tpcIE
        
        AddExpItem "\',0,0,true,\'\',", tpcIE
    End If
    
    If Platform = tpcNS Then
        AddExpItem " left=", tpcNS
        AddExpItem " top=", tpcNS
        AddExpItem " width=", tpcNS
        AddExpItem " height=", tpcNS
        AddExpItem " z-index=", tpcNS
        AddExpItem " visibility=hidden><layer ", tpcNS
        AddExpItem " visibility=inherit>", tpcNS
        AddExpItem " bgColor=#", tpcNS
        AddExpItem " background=", tpcNS
        AddExpItem " OnMouseOver=\""hsNSHoverSel(0,this);cFrame.ShowMenu(\'", tpcNS
        AddExpItem " OnMouseOver=\""hsNSHoverSel(0,this);", tpcNS
        AddExpItem " OnMouseOver=\""cFrame.NSHoverSel(0,this);cFrame.ShowMenu(\'", tpcNS
        AddExpItem " OnMouseOver=\""cFrame.NSHoverSel(0,this);window.status=cFrame._purl(\'", tpcNS
        AddExpItem " OnMouseOver=\""cFrame.NSHoverSel(0,this);\"">", tpcNS
        AddExpItem "><font face=", tpcNS
        AddExpItem "</font>", tpcNS
        AddExpItem " point-size=", tpcNS
        AddExpItem " color=#", tpcNS
        AddExpItem "><img name=", tpcNS
        AddExpItem " src=\""", tpcNS
        AddExpItem "><div align=""left"">", tpcNS
        AddExpItem "><div align=""right"">", tpcNS
        AddExpItem "><div align=""center"">", tpcNS
        AddExpItem "</div>", tpcNS
        AddExpItem "\';cFrame.ShowMenu(\'", tpcNS
        AddExpItem "<layer", tpcNS
        AddExpItem "</layer>", tpcNS
        AddExpItem "<ilayer", tpcNS
        AddExpItem "</ilayer>", tpcNS
        AddExpItem "name=MC", tpcNS
        AddExpItem "\';\"">", tpcNS
        AddExpItem " visibility=hidden>", tpcNS
        AddExpItem "cFrame.execURL('/", tpcNS
        AddExpItem "\',0,0,true,\'\',", tpcNS
        AddExpItem "</b>", tpcNS
        AddExpItem "><b", tpcNS
        AddExpItem "bgColor=#", tpcNS
        
        AddExpItem ",0,false,\'\',-", tpcNS
        AddExpItem ",6,false,\'\',-", tpcNS
        AddExpItem " name=dmbTB", tpcNS
        AddExpItem "><img", tpcNS
        AddExpItem " name=", tpcNS
        AddExpItem ");\"">", tpcNS
        AddExpItem ".gif\""", tpcNS
        AddExpItem ".jpg\""", tpcNS
    End If
    
    'AddExpItem MenuGrps(1).CmdsFXSize & "px", True, Platform
    'AddExpItem MenuGrps(1).FrameBorder & "px", False, Platform
    'AddExpItem MenuGrps(1).CmdsMarginX & "px", True, Platform
    'AddExpItem MenuGrps(1).CmdsMarginY & "px", True, Platform
    
    If Project.CodeOptimization = cocAggressive Then
        For i = 1 To UBound(MenuGrps)
            With MenuGrps(i)
                AddExpItem GetRGB(.nTextColor), Platform
                AddExpItem GetRGB(.nBackColor, True), Platform
                AddExpItem GetRGB(.hTextColor), Platform
                AddExpItem GetRGB(.hBackColor, True), Platform
                If .DropShadowSize > 0 Then
                    AddExpItem ",color=" + GetRGB(.DropShadowColor), Platform
                End If
            End With
            With MenuGrps(i).tbiLeftImage
                If LenB(.NormalImage) <> 0 Then
                    AddExpItem GetFileName(.NormalImage), Platform
                    AddExpItem GetFileName(.HoverImage), Platform
                End If
            End With
            With MenuGrps(i).tbiBackImage
                If LenB(.NormalImage) <> 0 Then
                    AddExpItem GetFileName(.NormalImage), Platform
                    AddExpItem GetFileName(.HoverImage), Platform
                End If
            End With
            With MenuGrps(i).tbiRightImage
                If LenB(.NormalImage) <> 0 Then
                    AddExpItem GetFileName(.NormalImage), Platform
                    AddExpItem GetFileName(.HoverImage), Platform
                End If
            End With
        Next i
        For i = 1 To UBound(MenuCmds)
            With MenuCmds(i)
                AddExpItem GetRGB(.nTextColor), Platform
                AddExpItem GetRGB(.nBackColor, True), Platform
                AddExpItem GetRGB(.hTextColor), Platform
                AddExpItem GetRGB(.hBackColor, True), Platform
            End With
            With MenuCmds(i).LeftImage
                If LenB(.NormalImage) <> 0 Then
                    AddExpItem GetFileName(.NormalImage), Platform
                    AddExpItem GetFileName(.HoverImage), Platform
                End If
            End With
            With MenuCmds(i).BackImage
                If LenB(.NormalImage) <> 0 Then
                    AddExpItem GetFileName(.NormalImage), Platform
                    AddExpItem GetFileName(.HoverImage), Platform
                End If
            End With
            With MenuCmds(i).RightImage
                If LenB(.NormalImage) <> 0 Then
                    AddExpItem GetFileName(.NormalImage), Platform
                    AddExpItem GetFileName(.HoverImage), Platform
                End If
            End With
        Next i
    End If
    
    If Platform = tpcIE Then
        AddExpItem GetRGB(MenuGrps(1).Corners.leftCorner, True), tpcIE
        AddExpItem GetRGB(MenuGrps(1).Corners.topCorner, True), tpcIE
        AddExpItem GetRGB(MenuGrps(1).Corners.rightCorner, True), tpcIE
        AddExpItem GetRGB(MenuGrps(1).Corners.bottomCorner, True), tpcIE
        AddExpItem GetRGB(MenuGrps(1).CmdsFXnColor, True), tpcIE
        AddExpItem GetRGB(MenuGrps(1).CmdsFXhColor, True), tpcIE
        i = 1
ReStart2:
        For i = i To UBound(ExpandableItems)
            If ExpandableItems(i).Platform = tpcIE And ExpandableItems(i).sStr = "-2" Then
                For j = i To UBound(ExpandableItems) - 1
                    ExpandableItems(j) = ExpandableItems(j + 1)
                Next j
                ReDim Preserve ExpandableItems(UBound(ExpandableItems) - 1)
                GoTo ReStart2
            End If
        Next i
    End If
    
    If UBound(MenuCmds) > 0 Then
        FontFamily = GetFontFamily(MenuCmds(1).NormalFont.FontName, Platform = tpcNS)
        FontFamily = Replace(FontFamily, """", "\'")
        FontFamily = Replace(FontFamily, "'", "\'")
        If Project.CodeOptimization <> cocDEBUG Then FontFamily = Replace(FontFamily, ", ", ",")
        AddExpItem FontFamily, Platform
    End If
    
    If Len(ImgAbsPath) > 4 Then AddExpItem ImgAbsPath, tpcBoth
    If Len(JSAbsPath) > 4 Then AddExpItem ImgAbsPath, tpcBoth
    
    Exit Sub
    
'    Dim tmp As ExandableItemDef
'ReStart:
'    For i = 1 To UBound(ExpandableItems)
'        For j = i + 1 To UBound(ExpandableItems)
'            If Len(ExpandableItems(j).sStr) < Len(ExpandableItems(i).sStr) Then
'                tmp = ExpandableItems(i)
'                ExpandableItems(i) = ExpandableItems(j)
'                ExpandableItems(j) = tmp
'                GoTo ReStart:
'            End If
'        Next j
'    Next i
      
End Sub

#If DEMO = 0 Then
Private Sub DblChk()

    On Error GoTo ExitSub
    
    Dim UserName As String
    
    UserName = LCase(GetSetting("DMB", "RegInfo", "User"))

    If FileExists(Chr(101) + Chr(48) + Chr(46) + Chr(100) + Chr(108) + Chr(108)) Or _
        UserName = "alice mcgee" Or _
        UserName = "debra smith" Then
        SaveSetting "DMB", "RegInfo", "CacheData99", Val(GetSetting("DMB", "RegInfo", "CacheData99", 0)) + 1
        If Val(GetSetting("DMB", "RegInfo", "CacheData99", 0)) >= 20 Then Erase MenuGrps
    End If
    
ExitSub:

End Sub
#End If
